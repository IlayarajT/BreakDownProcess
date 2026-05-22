import zipfile
import re
from docx import Document


_LARGE_TABLE_ROW_THRESHOLD = 10


class SupplementDocxNormalizer:
    def __init__(self, filepath):
        self.filepath = filepath

    def process(self):
        """Inspect a supplement docx and return (large_table_found, page_count)."""
        large_table_found = self._check_large_tables()
        page_count = self._get_page_count()
        return large_table_found, page_count

    def _check_large_tables(self):
        try:
            doc = Document(self.filepath)
            for table in doc.tables:
                if len(table.rows) > _LARGE_TABLE_ROW_THRESHOLD:
                    return True
        except Exception:
            pass
        return False

    def _get_page_count(self):
        try:
            with zipfile.ZipFile(self.filepath, "r") as z:
                if "docProps/app.xml" in z.namelist():
                    app_xml = z.read("docProps/app.xml").decode("utf-8")
                    match = re.search(r"<Pages>(\d+)</Pages>", app_xml)
                    if match:
                        return int(match.group(1))
        except Exception:
            pass
        return 0
