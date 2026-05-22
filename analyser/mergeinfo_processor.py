import os
import re
import json
import xmltodict
from TransformXml import XmlTransform

class MergeInfoProcessor:
    def __init__(self, analyser_config):
        self.analyser_config = analyser_config
        
    def create_info_file(self, customer, file_dics, jid, aid, ms_no, unique_id, process_folder):
        file_info = os.path.join(process_folder, "MergeInfo.xml")
        long_meta = file_dics.get('long_metadata', [None])[0]

        if long_meta and os.path.exists(long_meta):
            return self._process_long_metadata_file(
                long_meta, file_info, file_dics, 
                customer, jid, aid, process_folder, unique_id
            )
        else:
            return self._handle_missing_metadata(
                file_dics, customer, jid, aid, process_folder, unique_id
            )

    def _process_long_metadata_file(self, long_meta, file_info, file_dics, customer, jid, aid, process_folder, unique_id):
        # Clean XML file
        with open(long_meta, "r", encoding="utf8") as f:
            content = re.sub(r"<\!--s1(.+?)-->", "", f.read())
            content = re.sub(r"<\!DOCTYPE(.+?)>", "", content)
        
        with open(long_meta, "w", encoding="utf8") as f:
            f.write(content)

        # Transform XML
        xmltrns = XmlTransform()
        xsl_path = os.path.join(os.path.dirname(__file__), "xsl", "metaToFileInfo.xsl")
        success, info_xml = xmltrns.transform_xml(long_meta, file_info, xsl_path)
        
        if not success:
            return False, None, None

        # Parse XML to JSON
        with open(info_xml, encoding="utf8") as xml_file:
            data_dict = xmltodict.parse(xml_file.read())
        
        return self._process_file_info(
            data_dict, file_dics, customer, jid, aid, 
            process_folder, unique_id, info_xml
        )

    def _process_file_info(self, data_dict, file_dics, customer, jid, aid, process_folder, unique_id, info_xml):
        file_list = data_dict['root']['ItemType']
        file_categories = {"true": {}, "false": {}, "none": {}}
        doc_types = self.analyser_config['DocTypes']

        # Categorize files
        items = [file_list] if isinstance(file_list, dict) else file_list
        for item in items:
            merge_flag = item['ItemDescription']['@merge'].lower()
            file_type = item['ItemDescription']['@order_type']
            file_name = item['FileName']
            
            if file_type in doc_types:
                if merge_flag in file_categories:
                    file_categories[merge_flag].setdefault(file_type, []).append(file_name)

        # Prepare merger folder
        merger_input = self._get_merger_path(customer, jid, aid)
        self._prepare_merger_folder(merger_input, process_folder, jid, aid, unique_id, info_xml)

        # Process document files
        doc_files = {os.path.basename(p): p for p in file_dics.get('doc', [])}
        normalizer_input = {}
        
        # Check if automatic processing is possible
        if self._requires_manual_processing(file_categories, doc_files):
            return self._manual_process(doc_files, merger_input, customer, normalizer_input)
        else:
            return self._automatic_process(
                file_categories, doc_files, merger_input, 
                customer, normalizer_input
            )

    def _requires_manual_processing(self, file_categories, doc_files):
        true_list = file_categories["true"]
        false_list = file_categories["false"]
        
        # Check main document requirements
        if not true_list.get("main") or len(true_list["main"]) != 1:
            return True
        
        # Verify all required files exist
        for file_type, files in true_list.items():
            for file_name in files:
                if file_name not in doc_files:
                    return True
        
        return bool(false_list)

    def _manual_process(self, doc_files, merger_input, customer, normalizer_input):
        for doc_file in doc_files.values():
            shutil.copy(doc_file, os.path.join(merger_input, os.path.basename(doc_file)))
        
        normalizer_input['selected'] = {
            os.path.basename(f): os.path.join(merger_input, os.path.basename(f)) 
            for f in doc_files.values()
        }
        normalizer_input.update({
            "customer": customer,
            "folder": os.path.basename(merger_input),
            "process": "mMerge"
        })
        return False, None, normalizer_input

    def _automatic_process(self, file_categories, doc_files, merger_input, customer, normalizer_input):
        # Copy files to be merged
        merge_order = self.analyser_config['MergeOrder']
        for _, file_type in sorted(merge_order.items()):
            for file_name in file_categories["true"].get(file_type, []):
                src = doc_files[file_name]
                dst = os.path.join(merger_input, file_name)
                shutil.copy(src, dst)
                normalizer_input.setdefault("selected", {})[file_name] = dst

        # Copy files to be removed
        for file_name in sum(file_categories["none"].values(), []):
            src = doc_files[file_name]
            dst = os.path.join(merger_input, file_name)
            shutil.copy(src, dst)
            normalizer_input.setdefault("removed", {})[file_name] = dst

        normalizer_input.update({
            "customer": customer,
            "folder": os.path.basename(merger_input),
            "process": "mMerge"
        })
        return True, None, normalizer_input

    def _get_merger_path(self, customer, jid, aid):
        return re.sub(
            r"\[(CUSTOMER|JID|AID)\]", 
            lambda m: {"CUSTOMER": customer, "JID": jid, "AID": aid}[m.group(1)],
            self.analyser_config['FOLDERS']['MERGER']
        )

    def _prepare_merger_folder(self, merger_path, process_folder, jid, aid, unique_id, info_xml=None):
        os.makedirs(merger_path, exist_ok=True)
        
        # Write unique ID file
        with open(os.path.join(merger_path, "unique_id.json"), "w") as f:
            f.write(f'"unique_id": "{unique_id}"')
        
        # Copy JSON file
        json_file = f"{jid}_{aid}.json"
        shutil.copy(
            os.path.join(process_folder, json_file),
            os.path.join(merger_path, json_file)
        )
        
        # Copy info XML if exists
        if info_xml:
            shutil.copy(
                info_xml,
                os.path.join(merger_path, os.path.basename(info_xml))
            )
            
    def _handle_missing_metadata(self, file_dics, customer, jid, aid, process_folder, unique_id):
        # Implementation for missing metadata case
        pass