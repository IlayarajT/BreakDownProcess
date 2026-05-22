"""
labelAuthorsNER.py — Token-Stream HTML Pipeline for AU paragraph labeling.

ARCHITECTURE (replaces run-level XML manipulation):
═══════════════════════════════════════════════════
  WHY NOT EDIT RUNS DIRECTLY?
    Word fragments one logical text span across many <w:r> runs due to
    spellcheck boundaries, language marks, rsid attributes, tracked changes,
    font runs, etc.  "Xiaoxin Liu" may be 5 runs.  Inserting a new run at
    the right position is brittle and fails silently.

  THE BETTER WAY — 6-step pipeline:
    1. extract_tokens_from_para_xml()
       Walk every child of <w:p> via lxml, collect (text, is_superscript) pairs.
       Handles: <w:r>, <w:ins>, <w:del>, <w:fldSimple>, <w:hyperlink>.
       Preserves non-text nodes (drawings, EQ objects) as pass-through tokens.

    2. tokens_to_html()
       Collapse token list into a clean HTML string:
         "Xiaoxin Liu<sup>1†*</sup>, Kexin Zhao<sup>2†</sup>, Jie Zhou<sup>3*</sup>"

    3. parse_authors_from_html()
       Split on <sup> boundaries; cross-reference JSON for degree-stripping
       and name disambiguation.  JSON author count is ground truth.

    4. insert_labels_into_html()
       Label placement rules (in priority order):
         a) After <sup>...</sup> if present → most common case
         b) Before separator (, ; and) if no superscript found
         c) At end of author name if single author and no separator
       Produces: "Name<sup>1*</sup>[AU1], Name2<sup>2</sup>[AU2]"

    5. labeled_html_to_para_xml()
       Rebuild <w:p> from scratch:
         - Preserve original <w:pPr> (paragraph properties / style) exactly
         - Write clean runs: normal text run | superscript run | red label run
         - Pass-through any non-text nodes (drawings, EQ) unchanged

    6. replace_para_in_docx()
       Swap the old <w:p> XML for the new one inside the DOCX zip.

INTERFACES (backward compatible with breakDownProcess.py):
    from labelAuthorsNER import LabelAuthor
    proc = LabelAuthor()
    labeled_html, aut_dict = proc.author_process(au_html_str)

    from labelAuthorsNER import label_au_paragraph_in_docx
    ok, labeled_html, aut_dict = label_au_paragraph_in_docx(docx_path, json_path, output_path)

INSTALLATION:
    pip install python-docx lxml fuzzywuzzy python-Levenshtein beautifulsoup4
"""

from __future__ import annotations

import json
import logging
import os
import re
import string
import time
import unicodedata
import zipfile
from copy import deepcopy
from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Dict, List, Optional, Tuple

# ── Logging ───────────────────────────────────────────────────────────────────
logger = logging.getLogger("labelAuthorsNER")

from lxml import etree

# ── Optional dependencies ────────────────────────────────────────────────────
try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from fuzzywuzzy import fuzz, process as fw_process
    FUZZYWUZZY_AVAILABLE = True
except ImportError:
    FUZZYWUZZY_AVAILABLE = False

try:
    from bs4 import BeautifulSoup, NavigableString
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False

# ── XML namespace ────────────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

# Tags that contain inline text content we want to flatten
_INLINE_TAGS = frozenset([
    f"{{{W}}}r",
    f"{{{W}}}ins",
    f"{{{W}}}del",
    f"{{{W}}}hyperlink",
    f"{{{W}}}fldSimple",
    f"{{{W}}}sdtContent",
])

# Tags that are non-text objects — pass through unchanged
_PASSTHROUGH_TAGS = frozenset([
    f"{{{W}}}drawing",
    f"{{{W}}}pict",
    f"{{{W}}}object",
    f"{{{W}}}altChunk",
    f"{{{W}}}bookmarkStart",
    f"{{{W}}}bookmarkEnd",
    f"{{{W}}}commentRangeStart",
    f"{{{W}}}commentRangeEnd",
    f"{{{W}}}proofErr",
])


# ── File I/O utilities (production stability) ────────────────────────────────

def wait_for_file_ready(filepath: str, timeout: float = 10.0) -> bool:
    """
    Wait until *filepath* can be opened in read+write mode, i.e. no other
    process (antivirus, Word COM, Windows Search indexer, OneDrive sync)
    holds an exclusive lock on it.

    Returns True if the file is ready, False if the timeout expires.
    """
    start = time.time()
    while time.time() - start < timeout:
        try:
            with open(filepath, "r+b"):
                return True
        except (PermissionError, OSError):
            time.sleep(0.5)
    logger.warning(f"[wait_for_file_ready] File still locked after {timeout}s: {filepath}")
    return False


def os_replace_with_retry(
    src: str, dst: str, max_attempts: int = 5, delay: float = 0.5
) -> None:
    """
    Retry wrapper around os.replace().

    On Windows, antivirus / indexer / OneDrive sync can briefly hold the
    destination file open, causing a PermissionError.  Retrying a few
    times with a short delay handles this reliably.
    """
    for attempt in range(1, max_attempts + 1):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt < max_attempts:
                logger.warning(
                    f"[os_replace_with_retry] PermissionError replacing "
                    f"{dst} (attempt {attempt}/{max_attempts}), retrying in {delay}s..."
                )
                time.sleep(delay)
            else:
                logger.error(
                    f"[os_replace_with_retry] Cannot replace {dst} — "
                    f"file locked after {max_attempts} attempts"
                )
                raise


# ════════════════════════════════════════════════════════════════════════════
# Data structures
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Token:
    """A single logical unit extracted from the AU paragraph XML."""
    text: str = ""
    is_superscript: bool = False
    is_passthrough: bool = False       # True for drawings, bookmarks, etc.
    passthrough_elem: object = None    # lxml element to copy verbatim
    bold: bool = False
    italic: bool = False


@dataclass
class AuthorEntity:
    index: int
    name: str
    superscript: str = ""
    is_corresponding: bool = False
    orcid: str = ""
    json_key: str = ""
    degrees: str = ""                  # degrees text trailing the author

    @property
    def label(self) -> str:
        return f"[AU{self.index}]"


# ════════════════════════════════════════════════════════════════════════════
# Step 1 — Token extraction from <w:p> XML
# ════════════════════════════════════════════════════════════════════════════

def _is_superscript_run(r_elem: etree._Element) -> bool:
    """Return True if this <w:r> has superscript formatting."""
    rpr = r_elem.find(f"{{{W}}}rPr")
    if rpr is None:
        return False
    va = rpr.find(f"{{{W}}}vertAlign")
    if va is not None:
        val = va.get(f"{{{W}}}val", "")
        if val == "superscript":
            return True
    return False


def _run_text(r_elem: etree._Element) -> str:
    """Concatenate all <w:t> and <w:delText> text in a run."""
    parts = []
    for tag in (f"{{{W}}}t", f"{{{W}}}delText"):
        for t in r_elem.iter(tag):
            parts.append(t.text or "")
    return "".join(parts)


def _is_bold(r_elem: etree._Element) -> bool:
    rpr = r_elem.find(f"{{{W}}}rPr")
    if rpr is None:
        return False
    return rpr.find(f"{{{W}}}b") is not None


def _is_italic(r_elem: etree._Element) -> bool:
    rpr = r_elem.find(f"{{{W}}}rPr")
    if rpr is None:
        return False
    return rpr.find(f"{{{W}}}i") is not None


def extract_tokens_from_para_xml(p_elem: etree._Element) -> List[Token]:
    """
    Walk every direct and inline-nested child of <w:p>, returning a flat
    list of Token objects.

    Handles:
      - <w:r> (normal and superscript runs)
      - <w:ins> (accept tracked insertions)
      - <w:del> (skip deleted text)
      - <w:hyperlink> (treat as normal run container)
      - <w:fldSimple> (field — take text content)
      - Drawings, bookmarks, proofErr etc. → pass-through tokens

    The result is immune to run fragmentation: ten runs spelling "Xiaoxin Liu"
    produce one logical run of text "Xiaoxin Liu" when they have the same
    superscript state.  (We merge adjacent same-state tokens in tokens_to_html.)
    """
    tokens: List[Token] = []

    def _walk(elem: etree._Element, inside_del: bool = False):
        for child in elem:
            tag = child.tag

            # ── Skip <w:pPr> — it's paragraph properties, not content ──
            if tag == f"{{{W}}}pPr":
                continue

            # ── Pass-through nodes ──
            if tag in _PASSTHROUGH_TAGS:
                tokens.append(Token(is_passthrough=True, passthrough_elem=child))
                continue

            # ── Deleted runs — skip content but keep structure (pass-through) ──
            if tag == f"{{{W}}}del":
                # Do NOT recurse into del — we skip deleted text
                continue

            # ── Inserted runs — recurse, treating content as normal ──
            if tag == f"{{{W}}}ins":
                _walk(child, inside_del=False)
                continue

            # ── Normal run ──
            if tag == f"{{{W}}}r":
                text = _run_text(child)
                if text:
                    tokens.append(Token(
                        text=text,
                        is_superscript=_is_superscript_run(child),
                        bold=_is_bold(child),
                        italic=_is_italic(child),
                    ))
                # Check for tab, break, etc. — treat as space
                if child.find(f"{{{W}}}tab") is not None:
                    tokens.append(Token(text=" "))
                continue

            # ── Containers: hyperlink, fldSimple, sdtContent ──
            if tag in (_INLINE_TAGS - {f"{{{W}}}r", f"{{{W}}}ins", f"{{{W}}}del"}):
                _walk(child, inside_del)
                continue

            # ── Anything else (mc:AlternateContent etc.) ──
            tokens.append(Token(is_passthrough=True, passthrough_elem=child))

    _walk(p_elem)
    return tokens


# ════════════════════════════════════════════════════════════════════════════
# Step 2 — Tokens → flat HTML
# ════════════════════════════════════════════════════════════════════════════

def tokens_to_html(tokens: List[Token]) -> str:
    """
    Collapse token list into a simple HTML string.

    Adjacent same-state tokens are merged:
      - Adjacent superscript tokens → single <sup>...</sup>
        (Word often splits "1†*" into <sup>1</sup><sup>†</sup><sup>*</sup>)
      - Adjacent non-superscript tokens → concatenated plain text
    Pass-through tokens are ignored (they'll be re-inserted later from
    the original XML).

    Whitespace-only superscript runs (Word formatting artifacts) are handled
    context-sensitively:
      - Between two real sup tokens  → kept and merged (e.g. "1, 2" → <sup>1, 2</sup>)
      - Standalone / between plain text → dropped silently (avoids ghost <sup> </sup>)
    """
    parts: List[str] = []
    current_sup_text: List[str] = []  # accumulates adjacent superscript text

    def _flush_sup():
        if current_sup_text:
            merged = "".join(current_sup_text)
            stripped = merged.rstrip()
            trailing_ws = merged[len(stripped):]
            if stripped:
                parts.append(f"<sup>{stripped}</sup>")
            if trailing_ws:
                parts.append(trailing_ws)
            current_sup_text.clear()

    for tok in tokens:
        if tok.is_passthrough:
            continue
        text = tok.text
        if not text:
            continue

        if tok.is_superscript:
            if not text.strip():
                # Whitespace-only sup — keep only if already inside a sup group
                # (e.g. <sup>1,</sup><sup> </sup><sup>2</sup> → <sup>1, 2</sup>)
                # If standalone (nothing accumulated yet), drop it silently.
                if current_sup_text:
                    current_sup_text.append(text)
                # else: isolated space sup between plain-text segments → drop
            else:
                # Normal sup content — always accumulate
                current_sup_text.append(text)
        else:
            _flush_sup()
            parts.append(text)

    _flush_sup()  # emit any trailing superscript
    return "".join(parts)

def tokens_to_html_old(tokens: List[Token]) -> str:
    """
    Collapse token list into a simple HTML string.

    Adjacent same-state tokens are merged:
      - Adjacent superscript tokens → single <sup>...</sup>
        (Word often splits "1†*" into <sup>1</sup><sup>†</sup><sup>*</sup>)
      - Adjacent non-superscript tokens → concatenated plain text
    Pass-through tokens are ignored (they'll be re-inserted later from
    the original XML).
    """
    parts: List[str] = []
    # Track current run state for merging
    current_sup_text: List[str] = []  # accumulates adjacent superscript text

    def _flush_sup():
        """Emit accumulated superscript text as a single <sup> block."""
        if current_sup_text:
            parts.append(f"<sup>{''.join(current_sup_text)}</sup>")
            current_sup_text.clear()

    for tok in tokens:
        if tok.is_passthrough:
            continue
        text = tok.text
        if not text:
            continue
        if tok.is_superscript:
            current_sup_text.append(text)
        else:
            _flush_sup()
            parts.append(text)

    _flush_sup()  # emit any trailing superscript
    return "".join(parts)


# ════════════════════════════════════════════════════════════════════════════
# Step 3 — Parse authors from HTML + JSON cross-reference
# ════════════════════════════════════════════════════════════════════════════

def _normalize(name: str) -> str:
    """Lowercase, strip diacritics, strip punctuation/spaces."""
    name = unicodedata.normalize("NFKD", name)
    name = "".join(c for c in name if not unicodedata.combining(c))
    name = name.lower()
    name = re.sub(r"[^a-z0-9]", "", name)
    return name


def _fuzzy_match(name: str, candidates: Dict[str, str],
                 threshold: int = 70) -> Optional[str]:
    """
    Return the key in `candidates` whose value best matches `name`,
    or None if no match above threshold.
    candidates: {json_key: display_name}
    """
    if not candidates:
        return None
    if FUZZYWUZZY_AVAILABLE:
        result = fw_process.extractOne(
            name, candidates, scorer=fuzz.token_sort_ratio
        )
        if result:
            _, score, key = result
            if score >= threshold:
                return key
    else:
        # Fallback: normalized substring match
        norm_name = _normalize(name)
        for key, cand in candidates.items():
            if _normalize(cand) in norm_name or norm_name in _normalize(cand):
                return key
    return None


def load_json_authors(json_path: str) -> Tuple[int, Dict[str, dict]]:
    """Load author count and author info dict from article JSON."""
    if not json_path or not os.path.exists(json_path):
        return 0, {}
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    authors = data.get("authors_info", {})
    # Normalize: fix swapped first/last names, strip prefixes like "Dr."
    for key, info in authors.items():
        fn = (info.get("first-name") or "").strip()
        ln = (info.get("last-name") or "").strip()
        # Strip title prefixes from name fields (e.g. "Dr. Erika S." in last-name)
        for prefix in ["Dr.", "Dr", "Prof.", "Prof", "Mr.", "Mrs.", "Ms.", "Miss"]:
            if fn.startswith(prefix + " "):
                fn = fn[len(prefix):].strip()
            if ln.startswith(prefix + " "):
                ln = ln[len(prefix):].strip()
        info["first-name"] = fn
        info["last-name"] = ln
    return len(authors), authors


# ════════════════════════════════════════════════════════════════════════════
# Validation
# ════════════════════════════════════════════════════════════════════════════

class ValidationSeverity(Enum):
    INFO = auto()
    WARNING = auto()
    ERROR = auto()


@dataclass
class ValidationIssue:
    severity: ValidationSeverity
    code: str
    message: str


def validate_authors(
    authors: List[AuthorEntity],
    json_authors: Dict[str, dict],
    expected_count: int,
) -> List[ValidationIssue]:
    """
    Validate parsed author list against JSON ground truth.

    Returns a list of issues found.  Callers can decide how to handle
    each severity level (log, flag for manual review, abort, etc.).

    Checks performed:
      - Count mismatch (parsed vs JSON expected)
      - Unmatched JSON authors (JSON name with no fuzzy match in parsed list)
      - Empty or suspiciously short author names
      - Duplicate author names in parsed list
    """
    issues: List[ValidationIssue] = []

    # ── Count mismatch ────────────────────────────────────────────────────
    if expected_count > 0 and len(authors) != expected_count:
        diff = abs(len(authors) - expected_count)
        sev = ValidationSeverity.ERROR if diff > 1 else ValidationSeverity.WARNING
        issues.append(ValidationIssue(
            severity=sev,
            code="COUNT_MISMATCH",
            message=(f"Parsed {len(authors)} authors but JSON expects "
                     f"{expected_count} (diff={diff})"),
        ))

    # ── Unmatched JSON authors ────────────────────────────────────────────
    matched_keys = {a.json_key for a in authors if a.json_key}
    for key, info in json_authors.items():
        if key not in matched_keys:
            fn = (info.get("first-name") or "").strip()
            ln = (info.get("last-name") or "").strip()
            display = f"{fn} {ln}".strip()
            issues.append(ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="UNMATCHED_JSON_AUTHOR",
                message=f"JSON author '{display}' (key={key}) not matched "
                        f"to any parsed author",
            ))

    # ── Empty / suspicious names ──────────────────────────────────────────
    for a in authors:
        if not a.name or not a.name.strip():
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="EMPTY_AUTHOR_NAME",
                message=f"AU{a.index} has empty name",
            ))
        elif len(a.name.strip()) < 2:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="SHORT_AUTHOR_NAME",
                message=f"AU{a.index} name is suspiciously short: "
                        f"'{a.name}'",
            ))

    # ── Duplicate names ───────────────────────────────────────────────────
    norm_names: Dict[str, int] = {}
    for a in authors:
        nn = _normalize(a.name)
        if nn in norm_names:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="DUPLICATE_AUTHOR_NAME",
                message=f"AU{a.index} '{a.name}' duplicates "
                        f"AU{norm_names[nn]}",
            ))
        else:
            norm_names[nn] = a.index

    return issues


# ── Common degree abbreviations ──────────────────────────────────────────
# SINGLE CANONICAL SOURCE — used by both _DEGREE_PATTERNS regex and
# _is_degree() in _strip_degrees_from_name().  Add new degrees HERE ONLY.
# Sourced from au_tag.pl degrees list + production experience
_CANONICAL_DEGREES = frozenset({
    # Bachelor's
    'BA', 'BS', 'BSc', 'BHSc', 'BEng', 'BEd', 'BSN', 'BN', 'BPharm', 'BDS',
    # Master's
    'MA', 'MS', 'MSc', 'MBA', 'MEd', 'MPhil', 'MPH', 'MPA', 'MN', 'MSN',
    'MTech', 'MHPE', 'MRes', 'MCh',
    # Doctoral
    'PhD', 'DPhil', 'EdD', 'DrPH', 'DSc', 'PsyD', 'DMin', 'DBA', 'DLitt',
    # Medical
    'MD', 'DO', 'MBBS', 'MBChB', 'MB', 'BChir',
    # Nursing / Allied Health
    'RN', 'LPN', 'NP', 'CNS', 'APRN', 'CNM', 'CRNA', 'DNP', 'DPT',
    'OTD', 'OT', 'OTR',
    # Fellowships — Cardiology, GI, Surgery, etc.
    'FRCPC', 'FAAD', 'FAAN', 'FRCP', 'FACS', 'FCCP', 'FACC',
    'FAHA', 'FESC', 'FHRS', 'AGAF', 'MASGE', 'FJGES', 'FASGE',
    'FRS', 'FRCS', 'FRSE', 'FMedSci', 'FRCOG', 'FRCR', 'FRCPath',
    'FAAOS', 'FACP', 'FACR', 'FCAP',
    # Law
    'JD', 'LLB', 'LLM', 'LLD',
    # Other professional
    'CPA', 'CFA', 'PE', 'CEng', 'OBE', 'CBE', 'MBE', 'KBE',
    'CHES', 'CCRN', 'CDE',
    # Honours
    'HonD',
})

# Degrees with known mixed-case canonical forms (not all-caps)
_KNOWN_MIXED_CASE_DEGREES = frozenset({
    'PhD', 'DrPH', 'BSc', 'BHSc', 'MSc', 'MPhil', 'DPhil',
    'EdD', 'MBChB', 'MBBS', 'MHPE', 'PsyD', 'BPharm', 'MTech',
})

# Special degrees with non-alpha chars (handled separately in matching)
_SPECIAL_DEGREES = frozenset({'OTR/L', 'PA-C'})

# Pre-built upper-case lookup for fast membership test
_CANONICAL_DEGREES_UPPER = frozenset(d.upper() for d in _CANONICAL_DEGREES)

# Build the regex from the canonical set (supports optional dots)
def _build_degree_regex() -> re.Pattern:
    """Build degree-matching regex from _CANONICAL_DEGREES."""
    # Start with dotted patterns that need special regex handling
    dotted = []
    plain = []
    for d in sorted(_CANONICAL_DEGREES, key=len, reverse=True):
        if len(d) <= 3 and d.isalpha():
            # Short degrees like BA, BS, MA — allow optional dots: B.A.
            dotted.append(r'\.?'.join(d) + r'\.?')
        else:
            plain.append(re.escape(d))
    # Add special degrees
    plain.append(r'OTR/L')
    plain.append(r'PA\-C')
    all_patterns = dotted + plain
    return re.compile(
        r'(?i)\b(?:' + '|'.join(all_patterns) + r')\b',
        re.VERBOSE,
    )

_DEGREE_PATTERNS = _build_degree_regex()


def _is_degree(word: str) -> bool:
    """
    Check whether a word token is a degree abbreviation.

    Uses _CANONICAL_DEGREES as the single source of truth.
    Case-sensitivity rules prevent stripping real names like "Ma", "Ed", "Ba":
      - All-caps (MD, MBBS, FACC) → always a degree
      - Known mixed-case (PhD, BSc, MSc) → degree
      - Capitalized 3+ chars (MPhil, BEng) → degree
      - 2-char mixed case (Ma, Ed, Do) → NOT a degree (real surname)
    """
    w = word.rstrip('.')
    if not w:
        return False
    # Special degrees with non-alpha chars
    if w in _SPECIAL_DEGREES:
        return True
    # Must be in canonical set (case-insensitive lookup)
    if w.upper() not in _CANONICAL_DEGREES_UPPER:
        return False
    # Case-sensitivity guard to protect real names
    if w == w.upper():
        return True  # all-caps like MD, MBBS, FACC
    if w in _KNOWN_MIXED_CASE_DEGREES:
        return True  # known mixed-case like PhD, BSc, MSc
    if len(w) >= 3 and w[0].isupper():
        return True  # Capitalized 3+ chars like BSc, MPhil
    return False

# ── Surname prefixes (from au_tag.pl line 23) ────────────────────────────
# Used to keep multi-word surnames together during name parsing
_SURNAME_PREFIXES = frozenset([
    'al', 'el', 'ap', 'ben', 'della', 'delle', "dell'", 'dalle',
    'de la', 'de los', 'de', 'dela', 'del',
    'da', 'di', 'du', 'do', 'dos', 'das',
    'la', 'le', 'lo', 'les',
    "d'", "l'", "o'",
    'st.', 'st', 'san', 'santa',
    'den', 'der', 'det',
    'von der', 'von', 'vom',
    'van den', 'van der', 'van de', 'van',
    'ten', 'ter',
    'bin', 'binti', 'ibn',
    'abu', 'abd',
])

# ── Generational suffixes (from au_tag.pl line 26) ───────────────────────
_GEN_SUFFIXES = frozenset([
    'Jr.', 'Jr', 'Sr.', 'Sr', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII',
])

# ── Role/title words to strip (from au_tag.pl line 25) ───────────────────
_ROLE_PATTERNS = re.compile(
    r'(?:Associate\s+)?Professors?|(?:Assistant\s+)?Professor|'
    r'Editors?\-?in\-?Chief|(?:Guest\s+)?Editors?|Associate\s+Editors?|'
    r'(?:Graduate\s+)?Research\s+Assistants?|'
    r'Managing\s+Editors?|Moderators?|Students?|'
    r'Doctoral\s+(?:Candidate|Student|Fellow)|Research\s+Fellow|'
    r'Postdoctoral\s+(?:Fellow|Researcher)|'
    r'Lecturer|Senior\s+Lecturer|Reader',
    re.IGNORECASE,
)

# ── Group/consortium author patterns (from au_tag.pl line 34) ─────────────
_GROUP_AUTHOR_PATTERNS = re.compile(
    r'(?:on\s+behalf\s+of\s+(?:the\s+)?|for\s+(?:the\s+)?)?'
    r'(?:the\s+)?'
    r'(?:\w+\s+){0,3}'
    r'(?:Collaborators?|Consortium|Network|Working\s+Group|'
    r'Study\s+Group|Investigators?|Committee)',
    re.IGNORECASE,
)

def _strip_degrees_from_name(name: str) -> str:
    """
    Remove degree suffixes from a name string.
    "Jia Qi Adam Bai, BSc" → "Jia Qi Adam Bai"
    "Chaocheng Liu MD, FRCPC, FAAD" → "Chaocheng Liu"
    "MBBS, MD, MSc, FACC" → ""  (entire string is degrees)
    "Rui Ma" → "Rui Ma"  (Ma is not a degree — mixed case)

    Uses _CANONICAL_DEGREES as the single source of truth via _is_degree().
    """
    # First: strip multi-word role/degree titles (from au_tag.pl line 25)
    m_role = _ROLE_PATTERNS.search(name)
    if m_role and m_role.start() > 0:
        before = name[:m_role.start()].rstrip(' ,;')
        if before:
            return before
    
    # First: check if the ENTIRE string is just degrees (no real name at all)
    # Split on commas/semicolons/spaces and check each token
    tokens = re.split(r'[,;\s]+', name.strip().rstrip('.'))
    tokens = [t for t in tokens if t]
    if tokens and all(_is_degree(t) for t in tokens):
        return ""
    
    # Handle dotted degree patterns like "Ph.D.", "M.P.H.", "B.S.", "OTR/L"
    # These need special matching before the word-boundary-based regex
    dotted_deg = re.search(
        r'[,;\s](?:Ph\.?D\.?|M\.?P\.?H\.?|B\.?S\.?|M\.?S\.?|M\.?D\.?|'
        r'D\.?O\.?|B\.?A\.?|M\.?A\.?|M\.?Ed\.?|Ed\.?D\.?|Dr\.?PH\.?|'
        r'OTR/L|PA\-C)(?=[\s,;.]|$)',
        name, re.IGNORECASE
    )
    if dotted_deg:
        txt = dotted_deg.group().lstrip(' ,;')
        # Guard: verify the matched text is actually a degree, not a surname
        # like "Ma", "Ed", "Ba" that happen to match the optional-dot patterns
        if _is_degree(txt) or '.' in txt:
            before = name[:dotted_deg.start()].rstrip(' ,;')
            if before:
                return before
    
    # Normal stripping: find the earliest degree after a separator
    deg_pat = re.compile(
        r'(?<=[,;\s])(?:'
        + '|'.join(re.escape(d) for d in sorted(_CANONICAL_DEGREES, key=len, reverse=True))
        + r')\.?(?=[\s,;.]|$)',
        re.IGNORECASE
    )
    for m in deg_pat.finditer(name):
        txt = m.group().rstrip('.')
        if _is_degree(txt):
            before = name[:m.start()].rstrip(" ,;")
            if before:
                return before
            else:
                return ""
    
    # Check if string starts with a degree (no preceding separator)
    if tokens and _is_degree(tokens[0]):
        return ""
    
    return name


def _split_inter_sup_token(token: str,
                            next_json_name: Optional[str]) -> Tuple[str, str]:
    """
    Given text between two superscripts, e.g. ' MBBS, MSc, Pravin Nanga',
    split into (degrees_text, next_author_name).

    If next_json_name is provided, use it to find the split point.
    Otherwise use heuristics (capitalised word after separator).

    Handles all separator forms:
      ", Name"           →  comma separator
      "; Name"           →  semicolon separator
      ", and Name"       →  Oxford-comma + "and"
      "and Name"         →  bare "and" (no preceding comma)
      " MBBS, Name"      →  degrees then comma-separated name
      ", Alison D. Cox"  →  name with middle name not in JSON ("Alison Cox")
    """
    if not token.strip():
        return "", ""

    if next_json_name:
        j_parts = next_json_name.split()

        # Strategy A: try full name (exact substring match)
        idx = token.lower().find(next_json_name.lower())
        if idx >= 0:
            return token[:idx], token[idx:].rstrip(' ,;&*')

        # Strategy A2: try swapped name order (handles JSON first/last swap)
        if len(j_parts) >= 2:
            swapped_name = " ".join(reversed(j_parts))
            idx = token.lower().find(swapped_name.lower())
            if idx >= 0:
                return token[:idx], token[idx:].rstrip(' ,;&*')

        # Strategy B: find last name, then walk backwards past any middle
        #             names to the first preceding separator boundary.
        #
        # SHORT-NAME GUARD: For names ≤3 chars (He, Li, Ma, Wu, Xu), we
        # require the match to be at a word boundary AND not inside a known
        # degree abbreviation.  E.g. "He" inside "MHPE" or "CHES" should
        # NOT match, but ", He " should.
        if len(j_parts) >= 2:
            last_name = j_parts[-1]
            first_name = j_parts[0]

            # Try both name parts as possible last-name anchors
            # (in case JSON has first/last swapped)
            for try_last, try_first in [(last_name, first_name), (first_name, last_name)]:
                if len(try_last) < 2:
                    continue  # Skip single-char matches

                # For short names, use word-boundary regex instead of .find()
                if len(try_last) <= 3:
                    # Match only at word boundaries, NOT inside degree abbrevs
                    boundary_pat = re.compile(
                        r'(?<=[,;&\s])' + re.escape(try_last) + r'(?=[\s,;&]|$)',
                        re.IGNORECASE
                    )
                    bm = boundary_pat.search(token)
                    if not bm:
                        continue
                    li = bm.start()
                    # Verify the match is NOT inside a degree token
                    # Extract the "word" surrounding the match position
                    surrounding = token[max(0, li-5):li+len(try_last)+5]
                    surrounding_tokens = re.split(r'[,;\s]+', surrounding.strip())
                    matched_in_degree = any(
                        _is_degree(st) and try_last.lower() in st.lower() and st.lower() != try_last.lower()
                        for st in surrounding_tokens
                    )
                    if matched_in_degree:
                        continue
                else:
                    li = token.lower().find(try_last.lower())
                    if li < 0:
                        continue

                pre = token[:li]
                sep_match = list(re.finditer(r'[,;&]|\band\b', pre))
                if sep_match:
                    name_start = sep_match[-1].end()
                    while name_start < len(pre) and pre[name_start] in ' \t':
                        name_start += 1
                else:
                    fi = pre.lower().find(try_first.lower())
                    if fi >= 0:
                        name_start = fi
                    else:
                        name_start = len(pre) - len(pre.lstrip(' \t,;&'))
                degrees_part = token[:name_start]
                name_part = token[name_start: li + len(try_last)].strip()
                # Validate: name_part should contain at least 2 words or match a known name part
                if name_part and (len(name_part.split()) >= 2 or
                                  _normalize(try_first) in _normalize(name_part)):
                    return degrees_part, name_part

        # Strategy C: find first name (try both parts for swapped names)
        # Same short-name guard as Strategy B
        if j_parts:
            for part in j_parts:
                if len(part) < 2:
                    continue
                # For short names, require word boundary match
                if len(part) <= 3:
                    boundary_pat = re.compile(
                        r'(?<=[,;&\s])' + re.escape(part) + r'(?=[\s,;&]|$)',
                        re.IGNORECASE
                    )
                    bm = boundary_pat.search(token)
                    if not bm:
                        continue
                    fi = bm.start()
                    # Reject if inside a degree abbreviation
                    surrounding = token[max(0, fi-5):fi+len(part)+5]
                    surrounding_tokens = re.split(r'[,;\s]+', surrounding.strip())
                    if any(_is_degree(st) and part.lower() in st.lower() and st.lower() != part.lower()
                           for st in surrounding_tokens):
                        continue
                else:
                    fi = token.lower().find(part.lower())
                    if fi < 0:
                        continue
                return token[:fi], token[fi:].rstrip(' ,;&*')

    # Heuristic 1: uppercase letter after a separator (handles , ; , and, and)
    m = re.search(r"(?:[,;&]\s*(?:and\s+)?|(?:^|\s)and\s+)([A-Z])", token)
    if m:
        idx = m.start(1)
        return token[:idx], token[idx:].rstrip()

    # Heuristic 2: last resort — first multi-word capitalised name
    m2 = re.search(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)', token)
    if m2:
        return token[:m2.start()], token[m2.start():].rstrip()

    return token, ""


def _parse_no_sup_authors(
    html: str,
    json_authors: Dict[str, dict],
) -> List[AuthorEntity]:
    """
    Fallback for AU lines with NO superscript blocks at all.

    Handles:
      "Noa Gay Rúa, María Rueda-Extremera & María Cantero-García*"
      "Julia M.L. Pangalangan1, MS, DrPH, Beth Frates 2, MD, ..."
      (numbers inline as plain text, not marked superscript)

    Strategy: use JSON author names to locate each author in the plain text,
    in order.  When a name variant is found, capture text up to the next
    separator so middle names or alternate spellings are included.
    """
    import unicodedata

    def _norm_str(s: str) -> str:
        s = unicodedata.normalize("NFKD", s)
        s = "".join(c for c in s if not unicodedata.combining(c))
        return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()

    # Strip HTML tags for plain-text search
    plain = re.sub(r"<[^>]+>", "", html)
    plain_norm = _norm_str(plain)

    # Build a character-level mapping: plain_norm[i] ↔ plain[j]
    # (both strings may differ in length due to multi-char → single-char collapses)
    # Simpler: use plain directly for extraction, plain_norm for matching
    # Re-index by splitting: find each author's start in plain_norm, then
    # extract the corresponding slice of plain up to the next separator.
    
    # Pre-split plain on separators to get candidate "cells"
    SEPS = re.compile(r"[,&]|\s+and\s+")
    cells = SEPS.split(plain)          # ["Noa Gay Rúa", " María Rueda-Extremera", ...]
    cells_norm = [_norm_str(c) for c in cells]

    json_keys_ordered = sorted(json_authors.keys(), key=lambda k: int(k))
    authors: List[AuthorEntity] = []
    used_cell_indices: set = set()

    for au_idx_1, jkey in enumerate(json_keys_ordered, 1):
        info = json_authors[jkey]
        fn = info.get("first-name", "").strip()
        mn = info.get("middle-name", "").strip()
        ln = info.get("last-name", "").strip()

        # Build name variants: most-specific first
        # Also include swapped order in case JSON has first/last swapped
        name_variants = []
        if fn and mn and ln:
            name_variants.append(f"{fn} {mn} {ln}")
        if fn and ln:
            name_variants.append(f"{fn} {ln}")
            name_variants.append(f"{ln} {fn}")  # swapped order
        if fn:
            name_variants.append(fn)
        if ln:
            name_variants.append(ln)

        found = False
        for variant in name_variants:
            vnorm = _norm_str(variant)
            # First try: exact cell match
            for ci, cnorm in enumerate(cells_norm):
                if ci in used_cell_indices:
                    continue
                if vnorm in cnorm or cnorm in vnorm:
                    used_cell_indices.add(ci)
                    name_in_doc = cells[ci].strip().rstrip('*').strip()
                    # Strip trailing digits/nbsp (affiliation numbers)
                    name_in_doc = re.sub(r'[\s\u00a0\d,;]*$', '', name_in_doc).strip()
                    # Strip Unicode superscript digits (¹²³ etc.)
                    name_in_doc = _strip_unicode_superscripts(name_in_doc)
                    # Strip degrees
                    name_in_doc = _strip_degrees_from_name(name_in_doc)
                    is_corresp = '*' in cells[ci] or info.get("corresponding-author", False)
                    orcid = info.get("orcid", "") or ""
                    authors.append(AuthorEntity(
                        index=au_idx_1,
                        name=name_in_doc if name_in_doc else variant,
                        superscript="",
                        is_corresponding=is_corresp,
                        orcid=orcid,
                        json_key=jkey,
                    ))
                    found = True
                    break
            if found:
                break

            # Second try: substring anywhere in plain_norm
            pos = plain_norm.find(vnorm)
            if pos >= 0:
                # Extract from original plain up to next separator
                # Find approximate position in plain text
                # Map plain_norm pos → plain pos (approximate, good enough)
                plain_chunk = plain[pos: pos + len(variant) + 20]
                sep_m = SEPS.search(plain_chunk)
                name_in_doc = plain_chunk[:sep_m.start()].strip() if sep_m else plain_chunk.strip()
                name_in_doc = name_in_doc.rstrip('*').strip()
                name_in_doc = re.sub(r'[\s\u00a0\d,;]*$', '', name_in_doc).strip()
                is_corresp = ('*' in plain[pos: pos + len(variant) + 5]
                              or info.get("corresponding-author", False))
                orcid = info.get("orcid", "") or ""
                authors.append(AuthorEntity(
                    index=au_idx_1,
                    name=name_in_doc if name_in_doc else variant,
                    superscript="",
                    is_corresponding=is_corresp,
                    orcid=orcid,
                    json_key=jkey,
                ))
                found = True
                break

        if not found:
            display = " ".join(p for p in [fn, ln] if p)
            authors.append(AuthorEntity(
                index=au_idx_1,
                name=display,
                superscript="",
                is_corresponding=info.get("corresponding-author", False),
                orcid=info.get("orcid", "") or "",
                json_key=jkey,
            ))

    return authors


def parse_authors_from_html(
    html: str,
    json_authors: Dict[str, dict],
    expected_count: int = 0,
) -> List[AuthorEntity]:
    """
    Parse the flat AU HTML into a list of AuthorEntity.

    Label placement rules (implemented in insert_labels_into_html):
      a) After <sup>...</sup> if present
      b) Before separator if no sup on this author
      c) At end of line for single/last author

    Special cases handled:
      • Article title prepended to AU line (e.g. BBI): first "author" that
        doesn't match any JSON name is skipped.
      • No superscripts at all (e.g. BCB_1436899): delegates to
        _parse_no_sup_authors() which uses JSON name matching.
      • Numbers/markers in plain runs (not marked superscript in XML):
        treated as no-sup and handled by JSON-guided fallback.
      • * inline for corresponding author (no sup block).
    """
    # Build JSON name lookup: json_key → full display name
    # Also handle swapped first/last names in JSON
    json_display: Dict[str, str] = {}
    for key, info in json_authors.items():
        fn = (info.get("first-name") or "").strip()
        mn = (info.get("middle-name") or "").strip()
        ln = (info.get("last-name") or "").strip()
        json_display[key] = " ".join(p for p in [fn, mn, ln] if p)

    # ── No-sup fallback ───────────────────────────────────────────────────────
    # If there are no <sup> blocks at all (or none that contain digits/letters
    # that could be affiliation markers), use JSON-guided plain-text parsing.
    sup_blocks = re.findall(r"<sup>[^<]*</sup>", html)
    has_real_sups = any(
        re.search(r"[a-zA-Z0-9]", re.sub(r"</?sup>", "", s))
        for s in sup_blocks
    )
    if not has_real_sups and json_authors and expected_count > 1:
        return _parse_no_sup_authors(html, json_authors)

    # ── Superscript-boundary parsing ──────────────────────────────────────────
    parts = re.split(r"(<sup>[^<]*</sup>)", html)

    authors: List[AuthorEntity] = []
    au_idx = 0
    used_json_keys: set = set()

    json_keys_ordered = sorted(json_display.keys(), key=lambda k: int(k))

    i = 0
    while i < len(parts):
        part = parts[i]

        if part.startswith("<sup>"):
            i += 1
            continue

        # Plain-text segment between sups
        text_seg = part

        has_sup = (i + 1 < len(parts)) and parts[i + 1].startswith("<sup>")
        sup_content = ""
        if has_sup:
            sup_content = re.sub(r"</?sup>", "", parts[i + 1])

        # JSON name for next author (used for degree-stripping)
        next_key_idx = len(authors)
        next_json_name: Optional[str] = None
        if next_key_idx < len(json_keys_ordered):
            next_key = json_keys_ordered[next_key_idx]
            next_json_name = json_display.get(next_key)

        # Split: trailing-of-prev-author / name-of-this-author
        if authors:
            degrees_text, name_text = _split_inter_sup_token(text_seg, next_json_name)
            authors[-1].degrees = degrees_text.strip(" ,;")
        else:
            degrees_text = ""
            name_text = text_seg

        # Strip NBSP, leading/trailing separators, * corresponding markers
        name_text = name_text.replace("\u00a0", " ")
        name_clean = re.sub(r"^[\s,;&]+|and\s+", "", name_text).strip()
        name_clean = re.sub(r"[\s,;&*]+$", "", name_clean).strip()
        # Strip "Dr." / "Prof." prefix from the docx text
        name_clean = re.sub(r"^(?:Dr\.|Dr|Prof\.|Prof|Mr\.|Mrs\.|Ms\.|Miss)\s+", "", name_clean).strip()
        # Strip "Authors:" / "Name:" prefix
        name_clean = re.sub(r'^(?:Name|Authors?)\s*:\s*', '', name_clean).strip()

        if not name_clean:
            i += 2 if has_sup else 1
            continue

        # ── Discard article-title prefix ────────────────────────────────────
        # If first "author" doesn't match any JSON name, it's a title prefix.
        # If it PARTIALLY matches (title + author name merged), extract just the name.
        if not authors and json_authors:
            name_clean_stripped = _strip_degrees_from_name(name_clean)
            # Also strip "Name:" / "Authors:" prefix for matching
            name_clean_stripped = re.sub(r'^(?:Name|Authors?)\s*:\s*', '', name_clean_stripped).strip()
            nc = _normalize(name_clean_stripped)
            
            any_match = False
            matched_dname = ""
            for jkey, dname in json_display.items():
                nd = _normalize(dname)
                if nc and nd and (nc == nd or nd == nc):
                    any_match = True
                    break
                if nc and nd and (nd in nc):
                    any_match = True
                    matched_dname = dname
                    break
                if nc and nd and (nc in nd):
                    any_match = True
                    break
                # Check by last-name + first-name separately (handles abbreviations)
                info = json_authors.get(jkey, {})
                fn = (info.get("first-name") or "").strip()
                ln = (info.get("last-name") or "").strip()
                if ln and fn and len(ln) > 2:
                    ln_norm = _normalize(ln)
                    fn_norm = _normalize(fn)
                    if ln_norm in nc and fn_norm[:3] in nc:
                        # Last name fully matches + first few chars of first name match
                        any_match = True
                        break
                # Also try swapped name order
                d_parts = dname.split()
                if len(d_parts) >= 2:
                    swapped = d_parts[-1] + " " + " ".join(d_parts[:-1])
                    ns = _normalize(swapped)
                    if nc and ns and (nc in ns or ns in nc):
                        any_match = True
                        break
                # Fuzzy match as fallback
                if FUZZYWUZZY_AVAILABLE and nc and nd:
                    score = fuzz.token_sort_ratio(name_clean_stripped, dname)
                    if score >= 70:
                        any_match = True
                        break
            
            if not any_match and has_sup:
                # Skip this segment — it's the article title or other non-author text
                i += 2
                continue
            
            # Check if title text is prepended to the author name
            # e.g. "Evidence Supporting EMA Drug Approvals (2020–2023): Maximilian Siebert"
            # If the text is much longer than any JSON name, try to extract just the name
            if any_match and matched_dname and has_sup:
                # Find where the matched author name starts in name_clean
                for dname in json_display.values():
                    dname_stripped = _strip_degrees_from_name(dname)
                    # Try direct match first
                    idx_found = name_clean.lower().find(dname_stripped.lower())
                    if idx_found < 0:
                        # Try swapped name order
                        dp = dname_stripped.split()
                        if len(dp) >= 2:
                            swapped_dn = " ".join(reversed(dp))
                            idx_found = name_clean.lower().find(swapped_dn.lower())
                    if idx_found < 0:
                        # Try normalized matching (strips diacritics)
                        nc_norm = _normalize(name_clean)
                        dn_norm = _normalize(dname_stripped)
                        if dn_norm and dn_norm in nc_norm:
                            # Find approximate position by matching character-by-character
                            pos = nc_norm.find(dn_norm)
                            if pos > 0:
                                # Map back to original: count non-stripped chars up to pos
                                # Approximate: find the first capital letter near that position
                                # that starts the author's first name
                                fn_parts = dname_stripped.split()
                                if fn_parts:
                                    for ci in range(len(name_clean)):
                                        remaining = name_clean[ci:]
                                        if remaining and remaining[0].isupper():
                                            rn = _normalize(remaining[:len(dname_stripped)+5])
                                            if dn_norm[:6] in rn:
                                                idx_found = ci
                                                break
                    if idx_found > 0 and idx_found > 3:
                        # There's significant text before the author name — it's a title prefix
                        name_clean = name_clean[idx_found:].strip()
                        name_clean = re.sub(r"^[\s,;:&]+", "", name_clean).strip()
                        break

        if has_sup:
            # Strip degrees from author name
            name_clean = _strip_degrees_from_name(name_clean)
            
            # Skip if name is empty after degree stripping (entire text was degrees)
            if not name_clean.strip():
                i += 2
                continue

            au_idx += 1
            is_corresp = bool(re.search(r"[*†‡]", sup_content))

            # Cross-ref with JSON (positional first, fuzzy fallback)
            # Also try swapped name order for matching
            orcid = ""
            matched_key = ""
            if au_idx <= len(json_keys_ordered):
                cand_key = json_keys_ordered[au_idx - 1]
                cand_name = json_display.get(cand_key, "")
                norm_parsed = _normalize(name_clean)
                norm_cand = _normalize(cand_name)
                # Also try swapped order of JSON name
                cand_parts = cand_name.split()
                norm_swapped = _normalize(" ".join(reversed(cand_parts))) if len(cand_parts) >= 2 else ""
                if norm_parsed and norm_cand and (
                    norm_parsed in norm_cand or norm_cand in norm_parsed or
                    (norm_swapped and (norm_parsed in norm_swapped or norm_swapped in norm_parsed)) or
                    (FUZZYWUZZY_AVAILABLE and
                     fuzz.token_sort_ratio(name_clean, cand_name) >= 60)
                ):
                    matched_key = cand_key
                    used_json_keys.add(cand_key)

            if not matched_key:
                remaining = {k: v for k, v in json_display.items()
                             if k not in used_json_keys}
                matched_key = _fuzzy_match(name_clean, remaining) or ""
                if matched_key:
                    used_json_keys.add(matched_key)

            if matched_key and matched_key in json_authors:
                orcid = json_authors[matched_key].get("orcid", "") or ""
                if json_authors[matched_key].get("corresponding-author"):
                    is_corresp = True

            authors.append(AuthorEntity(
                index=au_idx,
                name=name_clean,
                superscript=sup_content,
                is_corresponding=is_corresp,
                orcid=orcid,
                json_key=matched_key,
            ))
            i += 2
        else:
            # No superscript — trailing/sole author or no-sup doc
            if name_clean:
                # ── Multi-author trailing tail ─────────────────────────────
                # If there are several remaining JSON authors and the tail
                # contains separators (commas, semicolons, "and", "&"),
                # split it and process each piece as its own author.
                # This handles lines like:
                #   "Raul Pozos,1 Todd Cornish, Melissa Macias, Maria Soper, ..."
                # where only the first author has a superscript.
                remaining_count = len(json_keys_ordered) - len(authors)
                tail_seps = re.search(r"[,;&]|\s+and\s+", name_clean)
                if remaining_count > 1 and tail_seps:
                    # Split by separators, keep non-empty trimmed pieces
                    pieces = re.split(r"[,;&]|\s+and\s+", name_clean)
                    pieces = [p.strip(" \u00a0,;&*") for p in pieces if p and p.strip()]
                    # Drop pieces that are pure degrees (e.g. "MD", "PhD")
                    cleaned_pieces = []
                    for p in pieces:
                        p_no_deg = _strip_degrees_from_name(p).strip()
                        if not p_no_deg:
                            continue
                        # Skip pieces that look like degree fragments only
                        if re.match(r'^[\s,;.]*$', p_no_deg):
                            continue
                        cleaned_pieces.append(p_no_deg)

                    if len(cleaned_pieces) > 1:
                        for piece in cleaned_pieces:
                            piece = re.sub(
                                r'^(?:Dr\.|Dr|Prof\.|Prof|Mr\.|Mrs\.|Ms\.|Miss)\s+',
                                '', piece
                            ).strip()
                            piece = re.sub(r'^(?:Name|Authors?)\s*:\s*', '', piece).strip()
                            if not piece:
                                continue

                            au_idx += 1
                            orcid_p = ""
                            matched_key_p = ""
                            if au_idx <= len(json_keys_ordered):
                                cand_key_p = json_keys_ordered[au_idx - 1]
                                cand_name_p = json_display.get(cand_key_p, "")
                                norm_parsed_p = _normalize(piece)
                                norm_cand_p = _normalize(cand_name_p)
                                cand_parts_p = cand_name_p.split()
                                norm_swapped_p = _normalize(
                                    " ".join(reversed(cand_parts_p))
                                ) if len(cand_parts_p) >= 2 else ""
                                if norm_parsed_p and norm_cand_p and (
                                    norm_parsed_p in norm_cand_p
                                    or norm_cand_p in norm_parsed_p
                                    or (norm_swapped_p and (
                                        norm_parsed_p in norm_swapped_p
                                        or norm_swapped_p in norm_parsed_p
                                    ))
                                    or (FUZZYWUZZY_AVAILABLE and
                                        fuzz.token_sort_ratio(piece, cand_name_p) >= 60)
                                ):
                                    matched_key_p = cand_key_p
                                    used_json_keys.add(cand_key_p)

                            if not matched_key_p:
                                remaining = {k: v for k, v in json_display.items()
                                             if k not in used_json_keys}
                                matched_key_p = _fuzzy_match(piece, remaining) or ""
                                if matched_key_p:
                                    used_json_keys.add(matched_key_p)

                            is_corresp_p = False
                            if matched_key_p and matched_key_p in json_authors:
                                orcid_p = json_authors[matched_key_p].get("orcid", "") or ""
                                if json_authors[matched_key_p].get("corresponding-author"):
                                    is_corresp_p = True

                            piece_clean = _strip_degrees_from_name(piece).strip()
                            piece_clean = re.sub(r"[*]", "", piece_clean).strip()

                            authors.append(AuthorEntity(
                                index=au_idx,
                                name=piece_clean if piece_clean else piece,
                                superscript="",
                                is_corresponding=is_corresp_p,
                                orcid=orcid_p,
                                json_key=matched_key_p,
                            ))
                        i += 1
                        continue

                # ── Single trailing author (original behaviour) ───────────
                # Strip degrees from name
                name_clean_stripped = _strip_degrees_from_name(name_clean)
                
                # Check if the remaining text after degree stripping is empty or
                # only contains degree-like text — if so, it's trailing degrees
                # from the previous author, not a new author
                if not name_clean_stripped.strip() or (
                    authors and  # there are already parsed authors
                    len(name_clean_stripped.split()) <= 1 and
                    re.match(r'^[\s,;.]*$', name_clean_stripped)
                ):
                    i += 1
                    continue
                
                name_clean = name_clean_stripped
                # Strip "Name:" or "Authors:" prefix
                name_clean = re.sub(r'^(?:Name|Authors?)\s*:\s*', '', name_clean).strip()
                
                au_idx += 1

                orcid = ""
                matched_key = ""
                if au_idx <= len(json_keys_ordered):
                    cand_key = json_keys_ordered[au_idx - 1]
                    matched_key = cand_key
                    used_json_keys.add(cand_key)
                    orcid = json_authors.get(cand_key, {}).get("orcid", "") or ""

                is_corresp = bool(re.search(r"[*]", name_clean))
                name_clean = re.sub(r"[*]", "", name_clean).strip()

                # For single-author documents, strip affiliation text after the name
                # e.g. "Kevan Harris, University of California, Los Angeles, USA"
                # Use JSON name to find where the actual name ends
                if matched_key and expected_count == 1:
                    cand_name = json_display.get(matched_key, "")
                    if cand_name:
                        # Try to find the JSON name within name_clean
                        cand_stripped = _strip_degrees_from_name(cand_name)
                        nc_norm = _normalize(name_clean)
                        cd_norm = _normalize(cand_stripped)
                        if cd_norm and cd_norm in nc_norm and len(name_clean) > len(cand_stripped) + 5:
                            # The name_clean has extra text — use just the JSON name portion
                            # Find the JSON name in the original text
                            idx_match = name_clean.lower().find(cand_stripped.split()[0].lower())
                            if idx_match >= 0:
                                # Extract from first match to end of last-name
                                ln = (json_authors[matched_key].get("last-name") or "").strip()
                                if ln:
                                    ln_pos = name_clean.lower().find(ln.lower(), idx_match)
                                    if ln_pos >= 0:
                                        name_clean = name_clean[idx_match: ln_pos + len(ln)].strip()

                if matched_key and json_authors.get(matched_key, {}).get("corresponding-author"):
                    is_corresp = True

                authors.append(AuthorEntity(
                    index=au_idx,
                    name=name_clean,
                    superscript="",
                    is_corresponding=is_corresp,
                    orcid=orcid,
                    json_key=matched_key,
                ))
            i += 1

    return authors


# ════════════════════════════════════════════════════════════════════════════
# Step 4 — Insert [AUn] labels into HTML
# ════════════════════════════════════════════════════════════════════════════

def insert_labels_into_html(html: str, authors: List[AuthorEntity]) -> str:
    """
    Insert [AUn] labels into the flat HTML string.

    Label placement (in order of preference):
      a) After <sup>...</sup> + any trailing markers (*†‡)  → "Name<sup>1</sup>*[AU1]"
      b) Before the separator (, ; and) if no sup present    → "Name[AU1], Name2"
      c) At the very end for single / trailing author        → "Name[AU1]"

    For documents with NO superscripts at all, labels are inserted after
    each author's name using positional matching from the parsed authors list.
    """
    if not authors:
        return html

    # Check if this is a no-superscript document.
    # NOTE: a <sup> that contains only markers (* † ‡ § #) is NOT considered
    # a "real" superscript by parse_authors_from_html — those authors are
    # parsed via _parse_no_sup_authors with no sup info. We must use the
    # same definition here so the label-insertion path matches the parsing
    # path; otherwise a single corresponding-author marker like
    # "Author<sup>*</sup>" would force us into the sup-bound branch with
    # only one sup chunk for N authors.
    sup_blocks_chk = re.findall(r"<sup>[^<]*</sup>", html)
    has_any_sup = any(
        re.search(r"[a-zA-Z0-9]", re.sub(r"</?sup>", "", s))
        for s in sup_blocks_chk
    )

    if not has_any_sup:
        # ── No-sup label insertion: find each author name and insert label after it ──
        return _insert_labels_no_sup(html, authors)

    # ── Superscript-based label insertion ──
    parts = re.split(r"(<sup>[^<]*</sup>)", html)
    result: List[str] = []
    au_idx_seen = 0

    # Corresponding-author markers that may trail after </sup>
    # e.g. "Name<sup>1†</sup>*, Next" — the * belongs with the previous author
    _TRAILING_MARKERS = re.compile(r'^([*†‡§#]+)')

    i = 0
    while i < len(parts):
        chunk = parts[i]

        if chunk.startswith("<sup>"):
            result.append(chunk)
            if au_idx_seen < len(authors):
                a = authors[au_idx_seen]
                au_idx_seen += 1
                # Check if the NEXT text chunk starts with trailing markers
                # that belong to this author (e.g. "*, " or "*")
                if i + 1 < len(parts) and not parts[i + 1].startswith("<sup>"):
                    next_text = parts[i + 1]
                    m = _TRAILING_MARKERS.match(next_text)
                    if m:
                        # Emit the markers BEFORE the label
                        result.append(m.group(1))
                        # Remove consumed markers from the next chunk
                        parts[i + 1] = next_text[m.end():]
                result.append(f'<span class="aulabel">[AU{a.index}]</span>')
            i += 1
        else:
            next_is_sup = (i + 1 < len(parts)) and parts[i + 1].startswith("<sup>")

            if not next_is_sup and au_idx_seen < len(authors):
                # Text chunk without a following sup. Could be:
                #  (a) the truly-last/single trailing author, or
                #  (b) several remaining authors crammed together with no sups
                #      (e.g. "Raul,1 Todd, Melissa, Maria, ..." — only Raul has a sup)
                remaining = len(authors) - au_idx_seen

                if remaining > 1 and re.search(r"[,;&]|\s+and\s+", chunk):
                    # Split chunk by separators and place a label after
                    # each author-name segment.
                    seg_pattern = re.compile(r'([,;&]|\s+and\s+)')
                    segments = seg_pattern.split(chunk)
                    # segments alternates: [text, sep, text, sep, text, ...]
                    out: List[str] = []
                    for seg in segments:
                        if seg_pattern.fullmatch(seg or ""):
                            # It's a separator — keep as-is
                            out.append(seg)
                            continue
                        if not seg or not seg.strip():
                            out.append(seg)
                            continue
                        # Author-name segment: keep trailing markers attached,
                        # but place label before the marker if any (e.g. "*").
                        # Simple approach: append label at end of trimmed name,
                        # preserving the leading whitespace of the segment.
                        leading_ws = re.match(r'^\s*', seg).group(0)
                        trailing_ws = re.search(r'\s*$', seg).group(0)
                        core = seg[len(leading_ws): len(seg) - len(trailing_ws)] \
                            if trailing_ws else seg[len(leading_ws):]
                        # Move trailing markers (*, †, ‡) outside the label
                        m_end = re.search(r'([*†‡§#]+)\s*$', core)
                        if m_end:
                            core_no_marker = core[: m_end.start()]
                            markers = m_end.group(1)
                        else:
                            core_no_marker = core
                            markers = ""

                        if au_idx_seen < len(authors):
                            a = authors[au_idx_seen]
                            au_idx_seen += 1
                            out.append(
                                leading_ws
                                + core_no_marker
                                + markers
                                + f'<span class="aulabel">[AU{a.index}]</span>'
                                + trailing_ws
                            )
                        else:
                            out.append(seg)
                    result.append("".join(out))
                else:
                    # Truly the last/only trailing author — original behaviour.
                    text = chunk
                    m_end = re.search(r'([*†‡§#]+)\s*$', text)
                    if m_end:
                        result.append(text)
                    else:
                        result.append(text)
                    a = authors[au_idx_seen]
                    au_idx_seen += 1
                    result.append(f'<span class="aulabel">[AU{a.index}]</span>')
            else:
                result.append(chunk)
            i += 1

    while au_idx_seen < len(authors):
        a = authors[au_idx_seen]
        au_idx_seen += 1
        result.append(f'<span class="aulabel">[AU{a.index}]</span>')

    return "".join(result)


def _insert_labels_no_sup(html: str, authors: List[AuthorEntity]) -> str:
    """
    Insert labels into HTML that has no "real" superscript blocks (i.e.,
    no <sup> blocks containing alphanumeric content). The HTML may still
    contain marker-only <sup> blocks like <sup>*</sup> for corresponding
    authors, so we cannot assume html == plain.

    Strategy:
      1. Strip <sup> tags but keep their content to build a "plain-text"
         view used purely for searching author names.
      2. Build a position-map from plain-text index -> html index.
      3. Find each author's end position in plain, translate to html, and
         insert label there. For an author whose name is immediately
         followed by a <sup>...</sup> marker block in html (e.g. the
         corresponding author's "*"), the label is placed AFTER the marker
         so we don't break "Hyunjun Kim<sup>*</sup>" into
         "Hyunjun Kim[AUn]<sup>*</sup>".

    "Sebastian Koos, Elias Steinhilper and Marco Bitschnau"
      → "Sebastian Koos[AU1], Elias Steinhilper[AU2] and Marco Bitschnau[AU3]"

    "Jeongseon Park, Taejun An, Hyunjun Kim<sup>*</sup>"
      → "Jeongseon Park[AU1], Taejun An[AU2], Hyunjun Kim<sup>*</sup>[AU3]"
    """
    # Build plain text (no tags) and a map from plain index -> html index.
    plain_chars: List[str] = []
    plain_to_html: List[int] = []  # plain_to_html[plain_idx] = html_idx of that char
    in_tag = False
    for i, ch in enumerate(html):
        if ch == "<":
            in_tag = True
            continue
        if ch == ">":
            in_tag = False
            continue
        if in_tag:
            continue
        plain_chars.append(ch)
        plain_to_html.append(i)
    plain = "".join(plain_chars)

    def _plain_end_to_html_end(plain_end: int) -> int:
        """Convert exclusive plain end-position to exclusive html end-position.
        Then advance over any <sup>...</sup> block that immediately follows
        in the html (so labels go AFTER markers like <sup>*</sup>).
        """
        if plain_end <= 0:
            return 0
        if plain_end >= len(plain_to_html):
            html_end = len(html)
        else:
            html_end = plain_to_html[plain_end - 1] + 1
        m = re.match(r'\s*<sup>[^<]*</sup>', html[html_end:])
        if m:
            html_end += m.end()
        return html_end

    insertions: List[Tuple[int, str]] = []
    used_plain_positions: set = set()

    for a in authors:
        name = a.name
        if not name:
            continue
        label = f'<span class="aulabel">[AU{a.index}]</span>'

        search_start = max(used_plain_positions) + 1 if used_plain_positions else 0
        idx = plain.lower().find(name.lower(), search_start)
        if idx < 0:
            idx = plain.lower().find(name.lower())
            if idx < 0 or idx in used_plain_positions:
                continue

        plain_end = idx + len(name)
        used_plain_positions.add(idx)
        html_pos = _plain_end_to_html_end(plain_end)
        insertions.append((html_pos, label))

    if not insertions:
        # Fallback: append all labels at the end
        return html + "".join(
            f'<span class="aulabel">[AU{a.index}]</span>' for a in authors
        )

    insertions.sort(key=lambda x: x[0])

    result = []
    last_pos = 0
    for pos, label in insertions:
        result.append(html[last_pos:pos])
        result.append(label)
        last_pos = pos
    result.append(html[last_pos:])

    return "".join(result)


# ════════════════════════════════════════════════════════════════════════════
# Step 5 — Rebuild <w:p> XML from labeled HTML
# ════════════════════════════════════════════════════════════════════════════

def _make_run(text: str, superscript: bool = False,
              label: bool = False) -> etree._Element:
    """
    Create a <w:r> element.

    label=True  → red colour + aulabel character style
    superscript → vertAlign=superscript
    """
    r = etree.SubElement(etree.Element("root"), f"{{{W}}}r")
    # rPr
    rpr = etree.SubElement(r, f"{{{W}}}rPr")

    if label:
        # Character style reference
        rstyle = etree.SubElement(rpr, f"{{{W}}}rStyle")
        rstyle.set(f"{{{W}}}val", "aulabel")
        # Red colour
        color = etree.SubElement(rpr, f"{{{W}}}color")
        color.set(f"{{{W}}}val", "FF0000")

    if superscript:
        va = etree.SubElement(rpr, f"{{{W}}}vertAlign")
        va.set(f"{{{W}}}val", "superscript")

    if not label and not superscript:
        # Remove empty rPr
        r.remove(rpr)

    # <w:t>
    t = etree.SubElement(r, f"{{{W}}}t")
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return r


def labeled_html_to_para_xml(
    labeled_html: str,
    original_p_elem: etree._Element,
    passthrough_tokens: List[Token],
) -> etree._Element:
    """
    Build a new <w:p> element from labeled HTML.

    Preserves:
      - Original <w:pPr> (paragraph style, indentation, etc.) — copied exactly
      - Pass-through tokens (drawings, bookmarks) — re-appended at end

    New run structure:
      - Plain text runs
      - Superscript runs  (<sup>...</sup>)
      - Label runs        (<span class="aulabel">...</span>)
    """
    # Create new <w:p>
    new_p = etree.Element(f"{{{W}}}p")

    # Copy original <w:pPr> verbatim
    orig_pPr = original_p_elem.find(f"{{{W}}}pPr")
    if orig_pPr is not None:
        new_p.append(deepcopy(orig_pPr))

    # Tokenise the labeled HTML
    token_pattern = re.compile(
        r'(<sup>(?P<sup>[^<]*)</sup>)'
        r'|(<span class="aulabel">(?P<label>[^<]*)</span>)'
        r'|(?P<text>[^<]+)',
        re.DOTALL,
    )

    for m in token_pattern.finditer(labeled_html):
        if m.group("sup") is not None:
            run = _make_run(m.group("sup"), superscript=True)
            new_p.append(run)
        elif m.group("label") is not None:
            run = _make_run(m.group("label"), label=True)
            new_p.append(run)
        elif m.group("text") is not None:
            txt = m.group("text")
            if txt:
                run = _make_run(txt)
                new_p.append(run)

    # Re-append pass-through tokens (drawings, bookmarks, etc.)
    for tok in passthrough_tokens:
        if tok.passthrough_elem is not None:
            new_p.append(deepcopy(tok.passthrough_elem))

    return new_p


# ════════════════════════════════════════════════════════════════════════════
# Step 6 — Replace paragraph in DOCX zip
# ════════════════════════════════════════════════════════════════════════════

def _get_au_para_xml(docx_path: str) -> Optional[Tuple[str, etree._Element]]:
    """
    Extract the AU/AU0 paragraph from the DOCX document.xml.
    Returns (document_xml_text, au_p_element) or None if not found.
    
    When a single AU paragraph exists, returns it directly.
    """
    if not wait_for_file_ready(docx_path):
        logger.error(f"[_get_au_para_xml] File locked, cannot access: {docx_path}")
        return None

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml").decode("utf-8")

    root = etree.fromstring(doc_xml.encode("utf-8"))
    ns = {"w": W}

    # Find AU or AU0 paragraphs
    au_paras = []
    for xpath in [
        "//w:p[w:pPr/w:pStyle[@w:val='AU']]",
        "//w:p[w:pPr/w:pStyle[@w:val='AU0']]",
    ]:
        results = root.xpath(xpath, namespaces=ns)
        if results:
            au_paras = results
            break

    if not au_paras:
        return None

    return doc_xml, au_paras[0]


def _get_all_au_paras(docx_path: str) -> Optional[Tuple[str, List[etree._Element]]]:
    """
    Extract ALL AU/AU0 paragraphs from the DOCX document.xml.
    Returns (document_xml_text, [au_p_elements]) or None if not found.
    """
    if not wait_for_file_ready(docx_path):
        logger.error(f"[_get_all_au_paras] File locked, cannot access: {docx_path}")
        return None

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml").decode("utf-8")

    root = etree.fromstring(doc_xml.encode("utf-8"))
    ns = {"w": W}

    au_paras = []
    for xpath in [
        "//w:p[w:pPr/w:pStyle[@w:val='AU']]",
        "//w:p[w:pPr/w:pStyle[@w:val='AU0']]",
    ]:
        results = root.xpath(xpath, namespaces=ns)
        if results:
            au_paras = results
            break

    if not au_paras:
        return None

    return doc_xml, au_paras


def replace_para_in_docx(
    docx_path: str,
    new_p_elem: etree._Element,
    original_p_elem: etree._Element,
    output_path: str,
) -> bool:
    """
    Replace the original AU paragraph in the DOCX zip with the new one.
    Writes output to output_path (can be same as docx_path for in-place edit).
    
    NOTE: original_p_elem must be from the SAME parsed tree as the root it
    belongs to — the parent is found via getparent().
    Returns True on success.
    """
    parent = original_p_elem.getparent()
    if parent is None:
        return False

    idx = list(parent).index(original_p_elem)
    parent.remove(original_p_elem)
    parent.insert(idx, new_p_elem)

    # Walk up from parent to find the document root
    root = parent
    while root.getparent() is not None:
        root = root.getparent()

    new_doc_xml = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    ).decode("utf-8")

    # Wait for file to be released before repacking
    if not wait_for_file_ready(docx_path):
        logger.error(f"[replace_para_in_docx] File locked before repack: {docx_path}")
        return False

    # Repack DOCX
    tmp_path = output_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin, \
         zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_doc_xml.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))

    os_replace_with_retry(tmp_path, output_path)
    return True


def replace_multiple_paras_in_docx(
    docx_path: str,
    new_p_elem: etree._Element,
    original_p_elems: List[etree._Element],
    output_path: str,
) -> bool:
    """
    Replace multiple AU paragraphs with a single new paragraph.
    The new paragraph replaces the first original; the rest are removed.
    All elements must be from the same parsed tree.
    """
    if not original_p_elems:
        return False
    
    # Replace first paragraph
    first = original_p_elems[0]
    parent = first.getparent()
    if parent is None:
        return False
    
    idx = list(parent).index(first)
    parent.remove(first)
    parent.insert(idx, new_p_elem)
    
    # Remove remaining AU paragraphs
    for p in original_p_elems[1:]:
        p_parent = p.getparent()
        if p_parent is not None:
            p_parent.remove(p)
    
    # Walk to root
    root = parent
    while root.getparent() is not None:
        root = root.getparent()

    new_doc_xml = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    ).decode("utf-8")

    # Wait for file to be released before repacking
    if not wait_for_file_ready(docx_path):
        logger.error(f"[replace_multiple_paras_in_docx] File locked before repack: {docx_path}")
        return False

    tmp_path = output_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin, \
         zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_doc_xml.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))

    os_replace_with_retry(tmp_path, output_path)
    return True


# ── Unicode superscript digits ────────────────────────────────────────────
_UNICODE_SUP_MAP = str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹', '0123456789')
_UNICODE_SUP_CHARS = set('⁰¹²³⁴⁵⁶⁷⁸⁹')


def _strip_unicode_superscripts(name: str) -> str:
    """Strip trailing Unicode superscript digits from a name.
    'Birgitta Schiller¹' → 'Birgitta Schiller'
    """
    return name.rstrip(''.join(_UNICODE_SUP_CHARS)).rstrip(',; ')


def _para_text(p_elem: etree._Element) -> str:
    """Get plain text from a paragraph element."""
    parts = []
    for t in p_elem.iter(f"{{{W}}}t"):
        parts.append(t.text or "")
    return "".join(parts).strip()


def _select_best_au_paras(
    au_paras: List[etree._Element],
    json_authors: Dict[str, dict],
    expected_count: int,
) -> List[etree._Element]:
    """
    From multiple AU paragraphs, select the one(s) that contain the actual
    author line. Filters out:
      - Empty paragraphs
      - Cover letter text (e.g. "Dear Dr Smith:")
      - Individual author entries (single author per paragraph, duplicate of main line)
      - Paragraphs with affiliation text (e.g. "Name, M.D.: Mayo Clinic Arizona")
      - Contact-info / email paragraphs (e.g. "Name Email: foo@bar.com")
      - CRediT author contribution paragraphs (e.g. "Name: Conceptualization, ...")
    
    Strategy: score each paragraph by how many JSON author names it contains.
    Pick the paragraph(s) with the highest score. If a single paragraph
    contains most/all authors, use just that one. If authors are split across
    2-3 consecutive paragraphs, use those.
    """
    if len(au_paras) <= 1:
        return au_paras
    
    # ── Patterns for non-author paragraphs ────────────────────────────────
    # Contact / address blocks:  "Name, MD, Department of ..., Email: ..."
    _CONTACT_RE = re.compile(
        r'(?:Email|E-mail|e-mail)\s*:', re.IGNORECASE
    )
    # Email address anywhere in text
    _EMAIL_RE = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    # CRediT roles
    _CREDIT_RE = re.compile(
        r'(?:Conceptualization|Methodology|Investigation|Formal\s+Analysis|'
        r'Writing\s+[–—-]\s+(?:Original|Review)|Supervision|Data\s+Curation|'
        r'Funding\s+Acquisition|Project\s+Administration|Resources|'
        r'Software|Validation|Visualization)',
        re.IGNORECASE
    )
    # Institutional affiliation block: "Name, Dept of X, University of Y"
    _AFFIL_RE = re.compile(
        r'(?:Department|Dept\.?|Division|School|Faculty|Center|Centre|'
        r'Hospital|Clinic|Institute|University|College|Laboratory|Lab)\s+(?:of|for)',
        re.IGNORECASE
    )
    # Cover letter
    _COVER_RE = re.compile(r'(?i)^(?:dear\s|editor)', re.IGNORECASE)
    
    def _is_contact_or_metadata(text: str) -> bool:
        """Return True if paragraph looks like contact/email/CRediT, not an author list."""
        if _CONTACT_RE.search(text):
            return True
        if _CREDIT_RE.search(text):
            return True
        # Email-based detection: if the paragraph contains email addresses,
        # check whether it's a contact/correspondence block vs a real author list.
        emails = _EMAIL_RE.findall(text)
        if emails:
            # Count how many distinct JSON last-names appear
            distinct_names = 0
            for info in json_authors.values():
                ln = (info.get("last-name") or "").strip()
                if ln and len(ln) > 2 and ln.lower() in text.lower():
                    distinct_names += 1
            # Few names + emails → contact block
            if distinct_names <= 2:
                return True
            # Many emails relative to names → correspondence list, not author line
            # (real AU lines rarely have emails; contact blocks have ~1 email per name)
            if len(emails) >= 3 and len(emails) >= distinct_names * 0.5:
                return True
        return False
    
    # Build normalized JSON name variants for matching
    json_names_norm = set()
    for info in json_authors.values():
        fn = (info.get("first-name") or "").strip()
        ln = (info.get("last-name") or "").strip()
        if fn and ln:
            json_names_norm.add(_normalize(f"{fn} {ln}"))
            json_names_norm.add(_normalize(f"{ln} {fn}"))  # swapped
        if ln:
            json_names_norm.add(_normalize(ln))
    
    # Score each paragraph
    scored = []
    for i, p in enumerate(au_paras):
        text = _para_text(p)
        text_norm = _normalize(text)
        
        # Count how many author names appear in this paragraph
        matches = sum(1 for n in json_names_norm if len(n) > 3 and n in text_norm)
        
        # Detect non-author content
        is_cover = bool(_COVER_RE.search(text))
        is_empty = len(text.strip()) < 3
        is_contact = _is_contact_or_metadata(text)
        has_affiliation = bool(_AFFIL_RE.search(text))
        
        scored.append((i, p, text, matches, is_cover, is_empty,
                        has_affiliation, is_contact))
    
    # Find the best paragraph(s)
    max_score = max(s[3] for s in scored)
    
    if max_score == 0:
        # No matches at all — return all non-empty, non-cover paragraphs
        return [s[1] for s in scored
                if not s[4] and not s[5] and not s[7]]
    
    # If one paragraph has most authors, use it (+ adjacent if they also have authors)
    best_paras = []
    for s in scored:
        idx, p, text, matches, is_cover, is_empty, has_affiliation, is_contact = s
        if is_cover or is_empty:
            continue
        # Skip contact/CRediT paragraphs unless they have the MOST authors
        # (edge case: author list paragraph might also contain an email)
        if is_contact and matches < max_score:
            continue
        if has_affiliation and matches < max_score:
            continue
        if matches > 0:
            best_paras.append(p)
    
    # If we got more paragraphs than needed (duplicate entries), keep only the ones
    # that have multiple authors (comma/and-separated names, not just degrees)
    if len(best_paras) > 2 and expected_count > 0:
        multi_author_paras = []
        for p in best_paras:
            text = _para_text(p)
            # Count how many JSON last-names appear in this paragraph
            name_hits = 0
            for info in json_authors.values():
                ln = (info.get("last-name") or "").strip()
                if ln and len(ln) > 2 and ln.lower() in text.lower():
                    name_hits += 1
            if name_hits >= 2:
                # This paragraph contains multiple author names — it's a real author line
                multi_author_paras.append(p)
        if multi_author_paras:
            best_paras = multi_author_paras
    
    return best_paras if best_paras else [au_paras[0]]


# ════════════════════════════════════════════════════════════════════════════
# Main pipeline functions
# ════════════════════════════════════════════════════════════════════════════

def label_au_paragraph_in_docx(
    docx_path: str,
    json_path: Optional[str] = None,
    output_path: Optional[str] = None,
) -> Tuple[bool, Optional[str], Dict[int, str]]:
    """
    Full 6-step pipeline: label the AU paragraph in a DOCX file.

    Handles multiple AU paragraphs by merging tokens from all of them
    into a single flat HTML, then replacing all AU paragraphs with one
    labeled paragraph (preserving the <w:pPr> of the first).

    Args:
        docx_path:    Path to input .docx
        json_path:    Path to article .json (for author count / names / ORCID)
        output_path:  Where to write the labeled DOCX.
                      Defaults to same as docx_path (in-place).

    Returns:
        (success, labeled_html_str, aut_dict)
        aut_dict: {1: "Author Name", 2: "Author Name", ...}
    """
    if output_path is None:
        output_path = docx_path

    # ── Load JSON ──────────────────────────────────────────────────────────
    expected_count, json_authors = load_json_authors(json_path) if json_path else (0, {})

    # ── Extract ALL AU paragraphs ──────────────────────────────────────────
    result = _get_all_au_paras(docx_path)
    if result is None:
        logger.warning("[LabelAU] No AU/AU0 paragraph found in document")
        return False, None, {}

    doc_xml, au_paras = result
    logger.info(f"[LabelAU] Found {len(au_paras)} AU paragraph(s)")

    # ── Filter AU paragraphs: pick the best candidate(s) ──────────────────
    # When there are many AU paragraphs, some may be cover letter text,
    # individual author entries (duplicates), or empty. We need the paragraph(s)
    # that actually contain the comma/and-separated author line.
    if len(au_paras) > 1 and json_authors:
        au_paras = _select_best_au_paras(au_paras, json_authors, expected_count)
        logger.info(f"[LabelAU] After filtering: {len(au_paras)} AU paragraph(s)")

    # ── Step 1: Extract tokens from ALL AU paragraphs ──────────────────────
    all_tokens: List[Token] = []
    for pi, au_p in enumerate(au_paras):
        para_tokens = extract_tokens_from_para_xml(au_p)
        if pi > 0 and all_tokens:
            # Add a space separator between paragraphs (they were separate lines)
            all_tokens.append(Token(text=" "))
        all_tokens.extend(para_tokens)

    passthrough_tokens = [t for t in all_tokens if t.is_passthrough]

    # ── Step 2: Flatten to HTML ────────────────────────────────────────────
    html = tokens_to_html(all_tokens)
    logger.info(f"[LabelAU] Flat HTML: {html[:120]}...")

    if not html.strip():
        logger.warning("[LabelAU] Empty AU paragraph")
        return False, None, {}

    # ── Step 3: Parse authors ──────────────────────────────────────────────
    authors = parse_authors_from_html(html, json_authors, expected_count)
    logger.info(f"[LabelAU] Parsed {len(authors)} authors "
          f"(JSON expected: {expected_count})")
    for a in authors:
        logger.debug(f"  AU{a.index}: {repr(a.name)} sup={repr(a.superscript)} "
              f"orcid={bool(a.orcid)} key={a.json_key}")

    if not authors:
        return False, None, {}

    # ── Validation gate ───────────────────────────────────────────────────
    # Check parsed authors against JSON ground truth and log issues.
    # Errors are logged but do NOT block labeling — downstream can decide.
    if json_authors:
        issues = validate_authors(authors, json_authors, expected_count)
        for issue in issues:
            if issue.severity == ValidationSeverity.ERROR:
                logger.error(f"[LabelAU:VALIDATE] {issue.code}: {issue.message}")
            elif issue.severity == ValidationSeverity.WARNING:
                logger.warning(f"[LabelAU:VALIDATE] {issue.code}: {issue.message}")
            else:
                logger.info(f"[LabelAU:VALIDATE] {issue.code}: {issue.message}")

    # ── Step 4: Insert labels ──────────────────────────────────────────────
    labeled_html = insert_labels_into_html(html, authors)
    logger.info(f"[LabelAU] Labeled HTML: {labeled_html[:200]}...")

    # ── Step 5: Rebuild <w:p> (using pPr from first AU paragraph) ─────────
    new_p = labeled_html_to_para_xml(labeled_html, au_paras[0], passthrough_tokens)

    # ── Step 6: Write back ─────────────────────────────────────────────────
    if len(au_paras) == 1:
        ok = replace_para_in_docx(docx_path, new_p, au_paras[0], output_path)
    else:
        ok = replace_multiple_paras_in_docx(docx_path, new_p, au_paras, output_path)

    if not ok:
        logger.error("[LabelAU] Failed to replace paragraph in DOCX")
        return False, None, {}

    # ── Build aut_dict for downstream use ─────────────────────────────────
    aut_dict: Dict[int, str] = {a.index: a.name for a in authors}

    return True, labeled_html, aut_dict


def label_au_html_string(
    au_html: str,
    json_authors: Optional[Dict[str, dict]] = None,
    expected_count: int = 0,
) -> Tuple[str, Dict[int, str]]:
    """
    Label an AU paragraph given as an HTML string (no DOCX involved).
    Used by the old LabelAuthor.author_process() interface.

    Returns:
        (labeled_html, aut_dict)
    """
    if json_authors is None:
        json_authors = {}

    # Parse <sup> from the HTML string (already converted from DOCX XML by
    # the existing XSLT pipeline in breakDownProcess.get_docx_authors)
    authors = parse_authors_from_html(au_html, json_authors, expected_count)
    if not authors:
        return au_html, {}

    labeled_html = insert_labels_into_html(au_html, authors)
    aut_dict = {a.index: a.name for a in authors}
    return labeled_html, aut_dict


# ════════════════════════════════════════════════════════════════════════════
# Backward-compatible LabelAuthor class
# ════════════════════════════════════════════════════════════════════════════

class LabelAuthor:
    """
    Drop-in replacement for the original labelAuthors.LabelAuthor.

    Interface used by breakDownProcess.py:
        proc_auth = LabelAuthor()
        labeled_html, aut_dict = proc_auth.author_process(au_html_string)

    The au_html_string comes from the existing XSLT pipeline that converts
    the docx AU XML to HTML.  This class just handles the labeling step.
    """

    def __init__(self, json_authors: Optional[Dict[str, dict]] = None):
        self._json_authors = json_authors or {}

    def set_json_authors(self, json_authors: Dict[str, dict]):
        self._json_authors = json_authors

    def author_process(
        self, author_html: str
    ) -> Tuple[str, Dict[int, str]]:
        """
        Process HTML author paragraph → (labeled_html, author_dict).

        Args:
            author_html: HTML like
                "Xiaoxin Liu<sup>1†*</sup>, Kexin Zhao<sup>2†</sup>"

        Returns:
            labeled_html: same string with [AUn] labels inserted
            aut_dict:     {1: "Xiaoxin Liu", 2: "Kexin Zhao", ...}
        """
        return label_au_html_string(
            author_html,
            json_authors=self._json_authors,
            expected_count=len(self._json_authors),
        )


# ════════════════════════════════════════════════════════════════════════════
# ORCID label insertion (used by write_labeled_docx)
# ════════════════════════════════════════════════════════════════════════════

ORCID_INSERT_TEXT = "[INSERT ORCID iD LOGO]"


def insert_orcid_labels_in_docx(
    docx_path: str,
    json_path: str,
    output_path: Optional[str] = None,
    gq_number: int = 1,
) -> Tuple[bool, str]:
    """
    Insert ORCID placeholder labels into the AU paragraph of a DOCX.

    Must be called AFTER label_au_paragraph_in_docx() so the [AUn] labels
    are already present.  The ORCID placeholder is inserted immediately after
    the author's superscript block (same position as the [AUn] label, but
    before it).

    Returns (success, output_path).
    """
    if output_path is None:
        output_path = docx_path

    expected_count, json_authors = load_json_authors(json_path)

    # Collect authors with ORCID
    orcid_authors: Dict[str, str] = {}  # normalized_name → orcid
    for key, info in json_authors.items():
        orcid = (info.get("orcid") or "").strip()
        if not orcid:
            continue
        fn = info.get("first-name", "")
        mn = info.get("middle-name", "")
        ln = info.get("last-name", "")
        for name_variant in [
            f"{fn} {ln}",
            f"{ln} {fn}",
            f"{fn} {mn} {ln}",
        ]:
            if name_variant.strip():
                orcid_authors[_normalize(name_variant.strip())] = orcid

    if not orcid_authors:
        return True, output_path  # Nothing to do

    result = _get_au_para_xml(docx_path)
    if result is None:
        return False, output_path

    _doc_xml, au_p = result
    tokens = extract_tokens_from_para_xml(au_p)
    passthrough_tokens = [t for t in tokens if t.is_passthrough]
    html = tokens_to_html(tokens)

    # Parse the already-labeled HTML (contains [AUn] spans from previous step)
    # We need to re-insert ORCID placeholders
    # Rebuild from scratch with ORCID markers
    authors = parse_authors_from_html(
        # Strip existing [AUn] labels first
        re.sub(r'<span class="aulabel">\[AU\d+\]</span>', "", html),
        json_authors,
        expected_count,
    )

    # Build labeled HTML with ORCID
    def _insert_orcid_and_label(html_clean: str,
                                 authors_list: List[AuthorEntity]) -> str:
        parts = re.split(r"(<sup>[^<]*</sup>)", html_clean)
        result: List[str] = []
        au_idx_seen = 0
        _TRAILING_MARKERS = re.compile(r'^([*†‡§#]+)')

        i = 0
        while i < len(parts):
            chunk = parts[i]
            if chunk.startswith("<sup>"):
                result.append(chunk)
                if au_idx_seen < len(authors_list):
                    a = authors_list[au_idx_seen]
                    au_idx_seen += 1
                    # Consume trailing markers from next text chunk
                    if i + 1 < len(parts) and not parts[i + 1].startswith("<sup>"):
                        next_text = parts[i + 1]
                        m = _TRAILING_MARKERS.match(next_text)
                        if m:
                            result.append(m.group(1))
                            parts[i + 1] = next_text[m.end():]
                    # Insert ORCID placeholder if this author has ORCID
                    if a.orcid:
                        result.append(
                            f'<span class="orcid">{ORCID_INSERT_TEXT}</span>'
                        )
                    result.append(
                        f'<span class="aulabel">[AU{a.index}]</span>'
                    )
                i += 1
            else:
                next_is_sup = (i + 1 < len(parts)) and parts[i + 1].startswith("<sup>")
                if not next_is_sup and au_idx_seen < len(authors_list):
                    result.append(chunk)
                    a = authors_list[au_idx_seen]
                    au_idx_seen += 1
                    if a.orcid:
                        result.append(
                            f'<span class="orcid">{ORCID_INSERT_TEXT}</span>'
                        )
                    result.append(
                        f'<span class="aulabel">[AU{a.index}]</span>'
                    )
                else:
                    result.append(chunk)
                i += 1
        return "".join(result)

    html_clean = re.sub(r'<span class="aulabel">\[AU\d+\]</span>', "", html)
    labeled_html = _insert_orcid_and_label(html_clean, authors)

    # Rebuild <w:p> — handle <span class="orcid"> as special red run
    new_p = etree.Element(f"{{{W}}}p")
    orig_pPr = au_p.find(f"{{{W}}}pPr")
    if orig_pPr is not None:
        new_p.append(deepcopy(orig_pPr))

    tok_pattern = re.compile(
        r'(<sup>(?P<sup>[^<]*)</sup>)'
        r'|(<span class="aulabel">(?P<label>[^<]*)</span>)'
        r'|(<span class="orcid">(?P<orcid>[^<]*)</span>)'
        r'|(?P<text>[^<]+)',
        re.DOTALL,
    )
    for m in tok_pattern.finditer(labeled_html):
        if m.group("sup"):
            new_p.append(_make_run(m.group("sup"), superscript=True))
        elif m.group("label"):
            new_p.append(_make_run(m.group("label"), label=True))
        elif m.group("orcid"):
            new_p.append(_make_run(m.group("orcid"), label=True))
        elif m.group("text"):
            txt = m.group("text")
            if txt:
                new_p.append(_make_run(txt))

    for tok in passthrough_tokens:
        if tok.passthrough_elem is not None:
            new_p.append(deepcopy(tok.passthrough_elem))

    # Append GQ callout
    gq_r = _make_run(f"[GQ: {gq_number}]", label=True)
    new_p.append(gq_r)

    ok = replace_para_in_docx(docx_path, new_p, au_p, output_path)
    return ok, output_path


# ════════════════════════════════════════════════════════════════════════════
# CLI
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python labelAuthorsNER.py <docx_path> [json_path] [output_path]")
        sys.exit(1)

    docx_path = sys.argv[1]
    json_path = sys.argv[2] if len(sys.argv) > 2 else None
    output_path = sys.argv[3] if len(sys.argv) > 3 else docx_path

    print(f"Input:  {docx_path}")
    print(f"JSON:   {json_path}")
    print(f"Output: {output_path}")

    ok, labeled_html, aut_dict = label_au_paragraph_in_docx(
        docx_path, json_path, output_path
    )

    print(f"\nSuccess: {ok}")
    print(f"Authors: {aut_dict}")
    if labeled_html:
        print(f"Labeled HTML:\n  {labeled_html}")
