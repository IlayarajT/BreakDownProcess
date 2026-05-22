import xml.etree.ElementTree as ET

class ShortMetadataProcessor:
    def __init__(self, sage_connector):
        self.sage_connector = sage_connector

    def process(self, file_path, process_folder, customer):
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Extract journal ID
            journal_id = self._get_journal_id(root)
            article_id = root.find('./front/article-meta/article-id[@pub-id-type="publisher-id"]').text
            ms_no = root.find('./front/article-meta/article-id[@pub-id-type="manuscript"]').text
            
            # Get article details from Sage Smart
            info_generated, article_id, jrn_tla = self.sage_connector.get_article_details(article_id, ms_no, journal_id, process_folder)
            
            return info_generated, journal_id, article_id, ms_no
        except Exception as e:
            return False, None, None, None

    def _get_journal_id(self, root):
        journal_id = root.find('./front/journal-meta/journal-id[@journal-id-type="acronym"]')
        if journal_id is not None:
            return journal_id.text
        return root.find('./front/journal-meta/journal-id[@journal-id-type="publisher"]').text.upper()