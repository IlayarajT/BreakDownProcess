import os
import json
from CreateArticleInfo import GetArticleId

class SageSmartConnector:
    def __init__(self, config_folder):
        self.config_folder = config_folder
        self.journal_info = self._load_journal_info()
        
    def _load_journal_info(self):
        journals_file = os.path.join(self.config_folder, "SupportingFiles", "sageJournalInfo.json")
        with open(journals_file, "r") as stream:
            return json.load(stream)
            
    def get_article_details(self, article_id, ms_no, journal_id, process_folder):
        create_info = GetArticleId()
        info_generated, article_id, jrn_tla = create_info.smart_login(
            article_id, ms_no, journal_id, process_folder
        )
        return info_generated, article_id, jrn_tla
        
    def get_journal_tla(self, journal_id):
        return self.journal_info.get(journal_id, {}).get("journal-tla")