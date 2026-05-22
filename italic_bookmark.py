"""
italic_bookmark.py

Two operations on a .docx file, selectable via mode:

  MODE 1 — "apply_bookmark"
      Scans every direct-body paragraph. If ALL text runs are italic,
      wraps the paragraph with a uniquely named bookmark:
          italic_para_1, italic_para_2, italic_para_3, ...
      The counter continues from whatever italic_para_* bookmarks already
      exist in the document, so repeated runs never collide.

  MODE 2 — "apply_italic"
      Scans every direct-body paragraph. If it carries any bookmark whose
      name matches italic_para_<number>, applies <w:i/> + <w:iCs/> to
      every run (and to the paragraph-level pPr > rPr), then removes
      the bookmarkStart / bookmarkEnd pair.

Both modes overwrite the input file atomically.

Requires: lxml  (pip install lxml)

CLI:
    python italic_bookmark.py apply_bookmark  document.docx
    python italic_bookmark.py apply_italic    document.docx

API:
    from italic_bookmark import ItalicBookmarkProcessor
    ItalicBookmarkProcessor("doc.docx", mode="apply_bookmark").process()
    ItalicBookmarkProcessor("doc.docx", mode="apply_italic").process()
"""

import re
import os
import sys
import shutil
import tempfile
import zipfile
from lxml import etree


class ItalicBookmarkProcessor:
    """
    Marks fully-italic body paragraphs with auto-incrementing bookmarks
    named italic_para_<N>, or restores italic formatting from those
    bookmarks and removes them.

    Only <w:p> elements that are direct children of <w:body> are touched.
    Table cells, text boxes, headers, and footers are left unchanged.

    Attributes:
        input_path (str) : .docx file to process (overwritten in place).
        mode       (str) : 'apply_bookmark' or 'apply_italic'.
    """

    # Bookmark name prefix and pattern
    BM_PREFIX  = "italic_para_"
    BM_PATTERN = re.compile(r"^italic_para_(\d+)$")

    W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def __init__(self, input_path: str, mode: str):
        """
        Args:
            input_path : Path to the .docx file (will be overwritten).
            mode       : 'apply_bookmark' or 'apply_italic'.

        Raises:
            ValueError : If mode is not one of the two accepted values.
        """
        if mode not in ("apply_bookmark", "apply_italic"):
            raise ValueError(
                f"mode must be 'apply_bookmark' or 'apply_italic', got '{mode}'"
            )
        self.input_path = input_path
        self.mode       = mode

        # Internal state reset on each process() call
        self._files: dict[str, bytes] = {}
        self._names: list[str]        = []
        self._next_bm_numeric_id: int = 0   # next w:id for <w:bookmarkStart>
        self._next_bm_counter:    int = 1   # next N in italic_para_N
        self._processed: int          = 0
        self._skipped:   int          = 0

    # ------------------------------------------------------------------ #
    # Public                                                               #
    # ------------------------------------------------------------------ #

    def process(self) -> dict:
        """
        Run the selected mode and overwrite the input file in place.

        Returns:
            dict with keys: file, mode, processed, skipped.
        """
        print(f"File : {self.input_path}")
        print(f"Mode : {self.mode}\n")

        self._read_docx()
        self._init_counters()
        self._process_document_xml()
        self._overwrite_in_place()

        result = {
            "file"      : self.input_path,
            "mode"      : self.mode,
            "processed" : self._processed,
            "skipped"   : self._skipped,
        }
        self._print_summary(result)
        return result

    # ------------------------------------------------------------------ #
    # Private — I/O                                                        #
    # ------------------------------------------------------------------ #

    def _read_docx(self):
        with zipfile.ZipFile(self.input_path, "r") as zin:
            self._names = zin.namelist()
            self._files = {n: zin.read(n) for n in self._names}

    def _overwrite_in_place(self):
        """Write to a sibling temp file then atomically replace the original."""
        dir_name = os.path.dirname(os.path.abspath(self.input_path))
        fd, tmp = tempfile.mkstemp(dir=dir_name, suffix=".docx.tmp")
        try:
            os.close(fd)
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for name in self._names:
                    zout.writestr(name, self._files[name])
            shutil.move(tmp, self.input_path)
        except Exception:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise

    # ------------------------------------------------------------------ #
    # Private — counter initialisation                                     #
    # ------------------------------------------------------------------ #

    def _init_counters(self):
        """
        Scan document.xml to set:
          _next_bm_numeric_id  — one above the highest existing w:id
          _next_bm_counter     — one above the highest existing italic_para_N
        so new bookmarks never collide with existing ones.
        """
        root = etree.fromstring(self._files["word/document.xml"])

        numeric_ids = []
        bm_counters = []

        for bm in root.findall(f".//{{{self.W}}}bookmarkStart"):
            # Collect all numeric w:id values
            raw_id = bm.get(f"{{{self.W}}}id", "")
            if raw_id.lstrip("-").isdigit():
                numeric_ids.append(int(raw_id))

            # Collect existing italic_para_N suffixes
            name = bm.get(f"{{{self.W}}}name", "")
            m = self.BM_PATTERN.match(name)
            if m:
                bm_counters.append(int(m.group(1)))

        self._next_bm_numeric_id = (max(numeric_ids) + 1) if numeric_ids else 0
        self._next_bm_counter    = (max(bm_counters) + 1) if bm_counters else 1

        print(f"  Next bookmark w:id     : {self._next_bm_numeric_id}")
        print(f"  Next italic_para_N     : italic_para_{self._next_bm_counter}\n")

    def _alloc_numeric_id(self) -> int:
        bid = self._next_bm_numeric_id
        self._next_bm_numeric_id += 1
        return bid

    def _alloc_bm_name(self) -> str:
        name = f"{self.BM_PREFIX}{self._next_bm_counter}"
        self._next_bm_counter += 1
        return name

    # ------------------------------------------------------------------ #
    # Private — document processing                                        #
    # ------------------------------------------------------------------ #

    def _process_document_xml(self):
        parser = etree.XMLParser(remove_blank_text=False)
        root   = etree.fromstring(self._files["word/document.xml"], parser)

        body = root.find(f"{{{self.W}}}body")
        if body is None:
            print("  WARNING: <w:body> not found.")
            return

        for para in body.findall(f"{{{self.W}}}p"):
            if self.mode == "apply_bookmark":
                self._handle_apply_bookmark(para)
            else:
                self._handle_apply_italic(para)

        self._files["word/document.xml"] = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    # ------------------------------------------------------------------ #
    # Private — Mode 1: apply_bookmark                                     #
    # ------------------------------------------------------------------ #

    def _handle_apply_bookmark(self, para):
        """
        If every text-bearing run in *para* is italic and the paragraph
        has no existing italic_para_* bookmark, insert a new uniquely
        named bookmark around all runs.
        """
        # Skip if already has an italic_para_* bookmark
        if self._existing_italic_bm_name(para) is not None:
            self._skipped += 1
            return

        text_runs = self._text_runs(para)
        if not text_runs or not self._all_italic(text_runs):
            self._skipped += 1
            return

        bm_name   = self._alloc_bm_name()
        numeric_id = self._alloc_numeric_id()

        # Build elements
        bm_start = etree.Element(f"{{{self.W}}}bookmarkStart")
        bm_start.set(f"{{{self.W}}}id",   str(numeric_id))
        bm_start.set(f"{{{self.W}}}name", bm_name)

        bm_end = etree.Element(f"{{{self.W}}}bookmarkEnd")
        bm_end.set(f"{{{self.W}}}id", str(numeric_id))

        # Locate insertion points among para's direct children
        children  = list(para)
        first_idx = children.index(text_runs[0])
        last_idx  = children.index(text_runs[-1])

        # Insert end before inserting start to avoid index shift
        para.insert(last_idx + 1, bm_end)
        para.insert(first_idx,    bm_start)

        self._processed += 1
        snippet = self._para_text(para)[:60].replace("\n", " ")
        print(f"  [bookmark added] {bm_name}  \"{snippet}...\"")

    # ------------------------------------------------------------------ #
    # Private — Mode 2: apply_italic                                       #
    # ------------------------------------------------------------------ #

    def _handle_apply_italic(self, para):
        """
        If *para* carries any italic_para_<N> bookmark, apply italic to
        every run (and pPr > rPr), then remove the bookmark elements.
        """
        bm_start, bm_end, bm_name = self._find_italic_bm_elements(para)
        if bm_start is None:
            self._skipped += 1
            return

        # Apply italic to paragraph-level rPr
        pPr = para.find(f"{{{self.W}}}pPr")
        if pPr is not None:
            self._add_italic_to_rpr(pPr)

        # Apply italic to every run
        for run in para.findall(f"{{{self.W}}}r"):
            self._add_italic_to_rpr(run)

        # Remove bookmark elements
        para.remove(bm_start)
        if bm_end is not None and bm_end in list(para):
            para.remove(bm_end)

        self._processed += 1
        snippet = self._para_text(para)[:60].replace("\n", " ")
        print(f"  [italic applied, bookmark removed] {bm_name}  \"{snippet}...\"")

    # ------------------------------------------------------------------ #
    # Private — italic helpers                                             #
    # ------------------------------------------------------------------ #

    def _add_italic_to_rpr(self, parent_elem):
        """
        Ensure <w:i/> and <w:iCs/> exist in the <w:rPr> child of
        *parent_elem* (a <w:pPr> or <w:r>). Creates <w:rPr> if absent.
        """
        rPr = parent_elem.find(f"{{{self.W}}}rPr")
        if rPr is None:
            rPr = etree.SubElement(parent_elem, f"{{{self.W}}}rPr")
        if rPr.find(f"{{{self.W}}}i") is None:
            rPr.append(etree.Element(f"{{{self.W}}}i"))
        if rPr.find(f"{{{self.W}}}iCs") is None:
            rPr.append(etree.Element(f"{{{self.W}}}iCs"))

    def _all_italic(self, runs: list) -> bool:
        """Return True only when every run has <w:i/> in its <w:rPr>."""
        for run in runs:
            rPr = run.find(f"{{{self.W}}}rPr")
            if rPr is None or rPr.find(f"{{{self.W}}}i") is None:
                return False
        return True

    # ------------------------------------------------------------------ #
    # Private — bookmark query helpers                                     #
    # ------------------------------------------------------------------ #

    def _existing_italic_bm_name(self, para) -> str | None:
        """
        Return the bookmark name if *para* already has an italic_para_*
        bookmarkStart, else None.
        """
        for bm in para.findall(f"{{{self.W}}}bookmarkStart"):
            name = bm.get(f"{{{self.W}}}name", "")
            if self.BM_PATTERN.match(name):
                return name
        return None

    def _find_italic_bm_elements(self, para):
        """
        Return (bookmarkStart, bookmarkEnd, name) for the first
        italic_para_* bookmark inside *para*, or (None, None, None).
        """
        bm_start = None
        bm_name  = None
        for bm in para.findall(f"{{{self.W}}}bookmarkStart"):
            name = bm.get(f"{{{self.W}}}name", "")
            if self.BM_PATTERN.match(name):
                bm_start = bm
                bm_name  = name
                break

        if bm_start is None:
            return None, None, None

        bm_id  = bm_start.get(f"{{{self.W}}}id")
        bm_end = None
        for bm in para.findall(f"{{{self.W}}}bookmarkEnd"):
            if bm.get(f"{{{self.W}}}id") == bm_id:
                bm_end = bm
                break

        return bm_start, bm_end, bm_name

    # ------------------------------------------------------------------ #
    # Private — text helpers                                               #
    # ------------------------------------------------------------------ #

    def _text_runs(self, para) -> list:
        """Direct <w:r> children of *para* that contain non-empty text."""
        return [
            r for r in para.findall(f"{{{self.W}}}r")
            if (r.find(f"{{{self.W}}}t") is not None)
            and (r.find(f"{{{self.W}}}t").text or "").strip()
        ]

    def _para_text(self, para) -> str:
        return "".join(
            (t.text or "") for t in para.findall(f".//{{{self.W}}}t")
        )

    # ------------------------------------------------------------------ #
    # Private — summary                                                    #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _print_summary(result: dict):
        action = (
            "Bookmarks added              "
            if result["mode"] == "apply_bookmark"
            else "Paragraphs italicised + bm removed"
        )
        print(
            f"\n{'='*54}\n"
            f"  Mode      : {result['mode']}\n"
            f"  {action} : {result['processed']}\n"
            f"  Skipped   : {result['skipped']}\n"
            f"  File      : {result['file']}\n"
            f"{'='*54}"
        )


# --------------------------------------------------------------------------- #
# CLI                                                                          #
# --------------------------------------------------------------------------- #

