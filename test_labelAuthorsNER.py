"""
test_labelAuthorsNER.py — Regression test suite for AU paragraph labeling.

Run:
    python test_labelAuthorsNER.py
    python -m pytest test_labelAuthorsNER.py -v

Each test case uses the HTML intermediate representation (Step 2 output)
plus a JSON author dict, and verifies that parse_authors_from_html (Step 3)
produces the correct author list.  This catches regressions in the parsing
logic without needing actual DOCX files.
"""

import re
import unittest

from labelAuthorsNER import (
    AuthorEntity,
    ValidationSeverity,
    _is_degree,
    _normalize,
    _split_inter_sup_token,
    _strip_degrees_from_name,
    insert_labels_into_html,
    parse_authors_from_html,
    validate_authors,
)


# ════════════════════════════════════════════════════════════════════════════
# Test fixtures — real-world HTML snippets from production documents
# ════════════════════════════════════════════════════════════════════════════

class TestParseAuthorsWithSup(unittest.TestCase):
    """Documents with superscript affiliation markers (most common case)."""

    def test_standard_comma_separated(self):
        """Standard: comma-separated authors, each with <sup> affiliations."""
        html = ('Lingli Qiu<sup>1,2</sup>, Wenqiang Zhang<sup>1</sup>, '
                'Zhixin Tan<sup>1</sup>, and Xia Jiang<sup>1,8,9,*</sup>')
        json_au = {
            '1': {'first-name': 'Lingli', 'last-name': 'Qiu'},
            '2': {'first-name': 'Wenqiang', 'last-name': 'Zhang'},
            '3': {'first-name': 'Zhixin', 'last-name': 'Tan'},
            '4': {'first-name': 'Xia', 'last-name': 'Jiang',
                  'corresponding-author': True},
        }
        authors = parse_authors_from_html(html, json_au, 4)
        self.assertEqual(len(authors), 4)
        self.assertEqual(authors[0].name, 'Lingli Qiu')
        self.assertEqual(authors[3].name, 'Xia Jiang')
        self.assertTrue(authors[3].is_corresponding)

    def test_with_degrees_between_sups(self):
        """Degrees (MD, FRCPC) appear between superscript and next author."""
        html = ('Chaocheng Liu<sup>1</sup> MD, FRCPC, FAAD, '
                'Jia Qi Adam Bai<sup>2</sup>, BSc')
        json_au = {
            '1': {'first-name': 'Chaocheng', 'last-name': 'Liu'},
            '2': {'first-name': 'Jia Qi Adam', 'last-name': 'Bai'},
        }
        authors = parse_authors_from_html(html, json_au, 2)
        self.assertEqual(len(authors), 2)
        self.assertEqual(authors[0].name, 'Chaocheng Liu')
        self.assertEqual(authors[1].name, 'Jia Qi Adam Bai')

    def test_semicolon_separator(self):
        """Authors separated by semicolons instead of commas."""
        html = ('Anna Smith<sup>1</sup>; Bob Jones<sup>2</sup>; '
                'Carol Lee<sup>3</sup>')
        json_au = {
            '1': {'first-name': 'Anna', 'last-name': 'Smith'},
            '2': {'first-name': 'Bob', 'last-name': 'Jones'},
            '3': {'first-name': 'Carol', 'last-name': 'Lee'},
        }
        authors = parse_authors_from_html(html, json_au, 3)
        self.assertEqual(len(authors), 3)

    def test_single_author_with_sup(self):
        """Single author with superscript."""
        html = 'Kevan Harris<sup>1,*</sup>'
        json_au = {
            '1': {'first-name': 'Kevan', 'last-name': 'Harris',
                  'corresponding-author': True},
        }
        authors = parse_authors_from_html(html, json_au, 1)
        self.assertEqual(len(authors), 1)
        self.assertEqual(authors[0].name, 'Kevan Harris')

    def test_swapped_json_name_order(self):
        """JSON has last-name in first-name field and vice versa."""
        html = 'Xiaoxin Liu<sup>1</sup>, Kexin Zhao<sup>2</sup>'
        json_au = {
            '1': {'first-name': 'Liu', 'last-name': 'Xiaoxin'},
            '2': {'first-name': 'Zhao', 'last-name': 'Kexin'},
        }
        authors = parse_authors_from_html(html, json_au, 2)
        self.assertEqual(len(authors), 2)
        self.assertEqual(authors[0].name, 'Xiaoxin Liu')


class TestParseAuthorsNoSup(unittest.TestCase):
    """Documents with NO superscript blocks at all."""

    def test_comma_and_separator(self):
        """Plain text with comma + 'and' separator."""
        html = 'Sebastian Koos, Elias Steinhilper and Marco Bitschnau'
        json_au = {
            '1': {'first-name': 'Sebastian', 'last-name': 'Koos'},
            '2': {'first-name': 'Elias', 'last-name': 'Steinhilper'},
            '3': {'first-name': 'Marco', 'last-name': 'Bitschnau'},
        }
        authors = parse_authors_from_html(html, json_au, 3)
        self.assertEqual(len(authors), 3)
        self.assertEqual(authors[0].name, 'Sebastian Koos')
        self.assertEqual(authors[2].name, 'Marco Bitschnau')

    def test_ampersand_separator(self):
        """Authors separated by &."""
        html = 'María Rueda-Extremera & María Cantero-García'
        json_au = {
            '1': {'first-name': 'María', 'last-name': 'Rueda-Extremera'},
            '2': {'first-name': 'María', 'last-name': 'Cantero-García'},
        }
        authors = parse_authors_from_html(html, json_au, 2)
        self.assertEqual(len(authors), 2)


# ════════════════════════════════════════════════════════════════════════════
# Test degree stripping
# ════════════════════════════════════════════════════════════════════════════

class TestDegreeStripping(unittest.TestCase):

    def test_strip_trailing_degrees(self):
        self.assertEqual(_strip_degrees_from_name('Jia Qi Adam Bai, BSc'),
                         'Jia Qi Adam Bai')

    def test_strip_multiple_degrees(self):
        self.assertEqual(_strip_degrees_from_name('Chaocheng Liu MD, FRCPC, FAAD'),
                         'Chaocheng Liu')

    def test_all_degrees_returns_empty(self):
        self.assertEqual(_strip_degrees_from_name('MBBS, MD, MSc, FACC'), '')

    def test_preserves_name_like_ma(self):
        """'Ma' should NOT be stripped — it's a surname, not a degree."""
        self.assertEqual(_strip_degrees_from_name('Rui Ma'), 'Rui Ma')

    def test_preserves_name_like_he(self):
        """'He' should NOT be stripped — it's a surname."""
        self.assertEqual(_strip_degrees_from_name('Qiurong He'), 'Qiurong He')

    def test_dotted_phd(self):
        self.assertEqual(_strip_degrees_from_name('John Smith, Ph.D.'),
                         'John Smith')

    def test_role_stripping(self):
        self.assertEqual(
            _strip_degrees_from_name('Jane Doe Associate Professor'),
            'Jane Doe')


class TestIsDegree(unittest.TestCase):
    """Test the unified _is_degree function."""

    def test_allcaps_degrees(self):
        for d in ['MD', 'MBBS', 'FACC', 'FRCPC', 'DNP', 'JD']:
            self.assertTrue(_is_degree(d), f'{d} should be a degree')

    def test_mixed_case_degrees(self):
        for d in ['PhD', 'BSc', 'MSc', 'MPhil', 'DPhil', 'EdD']:
            self.assertTrue(_is_degree(d), f'{d} should be a degree')

    def test_short_names_not_degrees(self):
        """2-char mixed-case words that are real names, not degrees."""
        for name in ['Ma', 'Ed', 'Ba', 'Do', 'He', 'Li']:
            self.assertFalse(_is_degree(name),
                             f'{name} should NOT be a degree')

    def test_special_degrees(self):
        self.assertTrue(_is_degree('OTR/L'))
        self.assertTrue(_is_degree('PA-C'))

    def test_dotted_stripped(self):
        """Trailing dots should be stripped before checking."""
        self.assertTrue(_is_degree('MD.'))
        self.assertTrue(_is_degree('PhD.'))


# ════════════════════════════════════════════════════════════════════════════
# Test short-name protection in _split_inter_sup_token
# ════════════════════════════════════════════════════════════════════════════

class TestSplitInterSupToken(unittest.TestCase):

    def test_short_name_not_matched_inside_degree(self):
        """'He' inside 'MHPE' should NOT match as author name 'He'."""
        token = ' MHPE, CHES, Qiurong He'
        degrees, name = _split_inter_sup_token(token, 'Qiurong He')
        self.assertIn('Qiurong', name)
        self.assertNotIn('MHPE', name)

    def test_short_name_ma_not_matched_in_degree(self):
        """'Ma' inside 'MA' degree should NOT match as author name 'Ma'."""
        token = ', MA, MSc, Rui Ma'
        degrees, name = _split_inter_sup_token(token, 'Rui Ma')
        self.assertIn('Rui', name)

    def test_normal_name_still_found(self):
        """Standard longer name should still be found correctly."""
        token = ', Pravin Nanga'
        degrees, name = _split_inter_sup_token(token, 'Pravin Nanga')
        self.assertEqual(name.strip(), 'Pravin Nanga')

    def test_degrees_then_name(self):
        """Degrees followed by author name."""
        token = ' MBBS, MSc, Pravin Nanga'
        degrees, name = _split_inter_sup_token(token, 'Pravin Nanga')
        self.assertIn('Pravin', name)
        self.assertIn('MBBS', degrees)

    def test_no_json_name_heuristic(self):
        """Without JSON hint, should fall back to heuristic (capital after sep)."""
        token = ', Anna Smith'
        degrees, name = _split_inter_sup_token(token, None)
        self.assertIn('Anna', name)


# ════════════════════════════════════════════════════════════════════════════
# Test label insertion
# ════════════════════════════════════════════════════════════════════════════

class TestInsertLabels(unittest.TestCase):

    def test_labels_after_sup(self):
        html = 'John Smith<sup>1</sup>, Jane Doe<sup>2</sup>'
        authors = [
            AuthorEntity(index=1, name='John Smith', superscript='1'),
            AuthorEntity(index=2, name='Jane Doe', superscript='2'),
        ]
        result = insert_labels_into_html(html, authors)
        self.assertIn('<sup>1</sup><span class="aulabel">[AU1]</span>', result)
        self.assertIn('<sup>2</sup><span class="aulabel">[AU2]</span>', result)

    def test_labels_after_trailing_star(self):
        """Label should come AFTER corresponding-author * marker."""
        html = 'Xiaoxin Liu<sup>1†</sup>*, Kexin Zhao<sup>2†</sup>, Jie Zhou<sup>3</sup>*'
        authors = [
            AuthorEntity(index=1, name='Xiaoxin Liu', superscript='1†',
                         is_corresponding=True),
            AuthorEntity(index=2, name='Kexin Zhao', superscript='2†'),
            AuthorEntity(index=3, name='Jie Zhou', superscript='3',
                         is_corresponding=True),
        ]
        result = insert_labels_into_html(html, authors)
        expected = ('Xiaoxin Liu<sup>1†</sup>*<span class="aulabel">[AU1]</span>'
                    ', Kexin Zhao<sup>2†</sup><span class="aulabel">[AU2]</span>'
                    ', Jie Zhou<sup>3</sup>*<span class="aulabel">[AU3]</span>')
        self.assertEqual(result, expected)

    def test_labels_no_sup(self):
        html = 'John Smith, Jane Doe'
        authors = [
            AuthorEntity(index=1, name='John Smith'),
            AuthorEntity(index=2, name='Jane Doe'),
        ]
        result = insert_labels_into_html(html, authors)
        self.assertIn('[AU1]', result)
        self.assertIn('[AU2]', result)


# ════════════════════════════════════════════════════════════════════════════
# Test validation
# ════════════════════════════════════════════════════════════════════════════

class TestValidation(unittest.TestCase):

    def test_count_mismatch_warning(self):
        """Off-by-1 should produce WARNING, not ERROR."""
        authors = [AuthorEntity(index=1, name='John Smith', json_key='1')]
        json_au = {
            '1': {'first-name': 'John', 'last-name': 'Smith'},
            '2': {'first-name': 'Jane', 'last-name': 'Doe'},
        }
        issues = validate_authors(authors, json_au, 2)
        count_issues = [i for i in issues if i.code == 'COUNT_MISMATCH']
        self.assertEqual(len(count_issues), 1)
        self.assertEqual(count_issues[0].severity, ValidationSeverity.WARNING)

    def test_count_mismatch_error(self):
        """Off-by-2+ should produce ERROR."""
        authors = [AuthorEntity(index=1, name='John Smith', json_key='1')]
        json_au = {
            '1': {'first-name': 'John', 'last-name': 'Smith'},
            '2': {'first-name': 'Jane', 'last-name': 'Doe'},
            '3': {'first-name': 'Bob', 'last-name': 'Lee'},
        }
        issues = validate_authors(authors, json_au, 3)
        count_issues = [i for i in issues if i.code == 'COUNT_MISMATCH']
        self.assertEqual(len(count_issues), 1)
        self.assertEqual(count_issues[0].severity, ValidationSeverity.ERROR)

    def test_no_issues_when_perfect(self):
        """No issues when parsed count matches and all JSON keys matched."""
        authors = [
            AuthorEntity(index=1, name='John Smith', json_key='1'),
            AuthorEntity(index=2, name='Jane Doe', json_key='2'),
        ]
        json_au = {
            '1': {'first-name': 'John', 'last-name': 'Smith'},
            '2': {'first-name': 'Jane', 'last-name': 'Doe'},
        }
        issues = validate_authors(authors, json_au, 2)
        self.assertEqual(len(issues), 0)

    def test_unmatched_json_author(self):
        authors = [AuthorEntity(index=1, name='John Smith', json_key='1')]
        json_au = {
            '1': {'first-name': 'John', 'last-name': 'Smith'},
            '2': {'first-name': 'Jane', 'last-name': 'Doe'},
        }
        issues = validate_authors(authors, json_au, 2)
        unmatched = [i for i in issues if i.code == 'UNMATCHED_JSON_AUTHOR']
        self.assertEqual(len(unmatched), 1)
        self.assertIn('Jane Doe', unmatched[0].message)

    def test_empty_name_error(self):
        authors = [AuthorEntity(index=1, name='', json_key='1')]
        json_au = {'1': {'first-name': 'John', 'last-name': 'Smith'}}
        issues = validate_authors(authors, json_au, 1)
        empty = [i for i in issues if i.code == 'EMPTY_AUTHOR_NAME']
        self.assertEqual(len(empty), 1)
        self.assertEqual(empty[0].severity, ValidationSeverity.ERROR)

    def test_duplicate_name_warning(self):
        authors = [
            AuthorEntity(index=1, name='John Smith', json_key='1'),
            AuthorEntity(index=2, name='John Smith', json_key='2'),
        ]
        json_au = {
            '1': {'first-name': 'John', 'last-name': 'Smith'},
            '2': {'first-name': 'John', 'last-name': 'Smith'},
        }
        issues = validate_authors(authors, json_au, 2)
        dupes = [i for i in issues if i.code == 'DUPLICATE_AUTHOR_NAME']
        self.assertEqual(len(dupes), 1)


# ════════════════════════════════════════════════════════════════════════════
# Edge cases from production
# ════════════════════════════════════════════════════════════════════════════

class TestEdgeCases(unittest.TestCase):

    def test_chinese_short_surnames(self):
        """Chinese authors with 2-char surnames (He, Li, Ma, Wu, Xu)."""
        html = ('Qiurong He<sup>1</sup>, Rui Ma<sup>2</sup>, '
                'Wei Li<sup>3</sup>')
        json_au = {
            '1': {'first-name': 'Qiurong', 'last-name': 'He'},
            '2': {'first-name': 'Rui', 'last-name': 'Ma'},
            '3': {'first-name': 'Wei', 'last-name': 'Li'},
        }
        authors = parse_authors_from_html(html, json_au, 3)
        self.assertEqual(len(authors), 3)
        self.assertEqual(authors[0].name, 'Qiurong He')
        self.assertEqual(authors[1].name, 'Rui Ma')
        self.assertEqual(authors[2].name, 'Wei Li')

    def test_surname_prefix_van_der(self):
        """Multi-word surname with prefix."""
        html = 'Jan van der Berg<sup>1</sup>, Anna de la Cruz<sup>2</sup>'
        json_au = {
            '1': {'first-name': 'Jan', 'last-name': 'van der Berg'},
            '2': {'first-name': 'Anna', 'last-name': 'de la Cruz'},
        }
        authors = parse_authors_from_html(html, json_au, 2)
        self.assertEqual(len(authors), 2)

    def test_corresponding_star_no_sup(self):
        """Corresponding author marked with * but no superscript."""
        html = 'Alice Brown*, Bob White'
        json_au = {
            '1': {'first-name': 'Alice', 'last-name': 'Brown',
                  'corresponding-author': True},
            '2': {'first-name': 'Bob', 'last-name': 'White'},
        }
        authors = parse_authors_from_html(html, json_au, 2)
        self.assertEqual(len(authors), 2)


if __name__ == '__main__':
    unittest.main()
