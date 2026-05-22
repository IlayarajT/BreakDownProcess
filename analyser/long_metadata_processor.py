import xml.etree.ElementTree as ET

class LongMetadataProcessor:
    def __init__(self, sage_connector):
        self.sage_connector = sage_connector

    def process(self, file_path, process_folder, customer):
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            journal_id = root.find('.//journal_abbreviation').text.upper()
            ms_no = root.find('./article').attrib.get('ms_no')
            
            if journal_id == customer:
                journal_id = ms_no.split("-")[0]
            
            info_generated, article_id, jrn_tla = self.sage_connector.get_article_details(
                None, ms_no, journal_id, process_folder
            )
            
            return info_generated, journal_id, article_id, ms_no
        except Exception as e:
            return False, None, None, None