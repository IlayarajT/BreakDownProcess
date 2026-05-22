import json
import logging
import os
import pathlib
import re
import shutil
import tarfile
import xml.etree.ElementTree as ET
import zipfile
import time
import py7zr
import xmltodict
import yaml
from rarfile import is_rarfile, RarFile

from CreateArticleInfo import GetArticleId
from TransformXml import XmlTransform
from dbprocess import DataBase
from loadconfig import getconfig


OFFICE_ZIP_EXTS = {".docx", ".xlsx", ".pptx", ".odt"}


class mAnalyzer:
    def __init__(self):
        self.configFolder, self.breakDownConfig = getconfig()

        analyser_yaml = os.path.join(self.configFolder, "config", "mAnalyser.yaml")
        with open(analyser_yaml, "r") as stream:
            self.analyser_config = yaml.safe_load(stream)

        journals_file = os.path.join(
            self.configFolder, "SupportingFiles", "sageJournalInfo.json"
        )
        with open(journals_file, "r") as stream:
            self.journal_json = json.load(stream)

        self.db_process = DataBase()
        self.nlogger = logging.getLogger("mAnalyser")

        self.processFolder = re.sub(
            r"\\", "/", self.analyser_config["FOLDERS"]["PROCESS"]
        )
        self.errorFolder = re.sub(
            r"\\", "/", self.analyser_config["FOLDERS"]["ERROR"]
        )
        self.logFolder = re.sub(
            r"\\", "/", self.analyser_config["LOGGER"]["BREAK_DOWN"]
        )

        if not self.processFolder.endswith("/"):
            self.processFolder += "/"
        if not self.errorFolder.endswith("/"):
            self.errorFolder += "/"
        if not self.logFolder.endswith("/"):
            self.logFolder += "/"

    # --------------------------------------------------
    # PATH UTILS
    # --------------------------------------------------

    @staticmethod
    def slash_replace(filepath):
        return re.sub(r"\\", "/", filepath)

    # --------------------------------------------------
    # FILE LIST
    # --------------------------------------------------

    def get_file_list(self, process_folder):
        input_files = {
            "process_folder": process_folder,
            "short_metadata": [],
            "long_metadata": [],
            "manifest": [],
            "doc": [],
            "office": [],
            "graphics": [],
            "tex": [],
            "pdf": [],
            "merge_info": [],
            "unknown": [],
        }

        for path, _, files in os.walk(process_folder):
            for name in files:
                if name.startswith("~$"):
                    continue

                filename = self.slash_replace(os.path.join(path, name))
                ext = pathlib.Path(name).suffix.lower()

                if re.match(r"SAGE-metadata-.*\.xml", name) or re.match(r"\d{2}-\d+\.xml", name):
                    input_files["short_metadata"].append(filename)
                elif name.endswith("-metadata.xml"):
                    input_files["long_metadata"].append(filename)
                elif name.endswith("-manifest.html"):
                    input_files["manifest"].append(filename)
                elif ext in (".doc", ".docx"):
                    input_files["doc"].append(filename)
                elif ext in (".xls", ".xlsx", ".ppt", ".pptx"):
                    input_files["office"].append(filename)
                elif ext in (".tif", ".tiff", ".jpg", ".jpeg", ".png", ".svg", ".eps"):
                    input_files["graphics"].append(filename)
                elif ext == ".tex":
                    input_files["tex"].append(filename)
                elif ext == ".pdf":
                    input_files["pdf"].append(filename)
                elif re.search("MergeInfo", name, re.IGNORECASE):
                    input_files["merge_info"].append(filename)
                else:
                    input_files["unknown"].append(filename)

        return input_files

    # --------------------------------------------------
    # ARCHIVE EXTRACTION (DOCX SAFE)
    # --------------------------------------------------

    def _recurse_archives(self, root_folder):
        for root, _, files in os.walk(root_folder):
            for fname in files:
                fpath = os.path.join(root, fname)
                ext = pathlib.Path(fname).suffix.lower()

                if ext in OFFICE_ZIP_EXTS:
                    continue

                if ext == ".zip" and zipfile.is_zipfile(fpath):
                    self.extract_zip(fpath, root)
                elif ext in (".tar", ".tgz", ".tar.gz") and tarfile.is_tarfile(fpath):
                    self.extract_tar(fpath, root)
                elif ext == ".rar" and is_rarfile(fpath):
                    self.extract_rar(fpath, root)
                elif ext == ".7z" and self.is_sevenZfile(fpath):
                    self.extract_sevenz(fpath, root)

    def extract_zip(self, filepath, toFolder):
        with zipfile.ZipFile(filepath, "r") as zf:
            zf.extractall(toFolder)
        os.remove(filepath)
        self._recurse_archives(toFolder)

    def extract_tar(self, filepath, toFolder):
        with tarfile.open(filepath, "r") as tf:
            tf.extractall(toFolder)
        os.remove(filepath)
        self._recurse_archives(toFolder)

    def extract_rar(self, filepath, toFolder):
        with RarFile(filepath, "r") as rf:
            rf.extractall(toFolder)
        os.remove(filepath)
        self._recurse_archives(toFolder)

    def extract_sevenz(self, filepath, toFolder):
        with py7zr.SevenZipFile(filepath, "r") as z:
            z.extractall(toFolder)
        os.remove(filepath)
        self._recurse_archives(toFolder)

    def is_sevenZfile(self, file_path):
        try:
            with py7zr.SevenZipFile(file_path, "r"):
                return True
        except Exception:
            return False

    def my_rarfile(self, file_path):
        return is_rarfile(file_path)

    # --------------------------------------------------
    # MOVE TO PROCESS
    # --------------------------------------------------

    def move_to_process(self, filepath, unique_id):
        name = pathlib.Path(filepath).stem
        process_folder = self.processFolder + name

        try:
            if os.path.exists(process_folder):
                shutil.rmtree(process_folder)

            if os.path.isdir(filepath):
                shutil.move(filepath, process_folder)
                return True, process_folder

            if zipfile.is_zipfile(filepath):
                self.extract_zip(filepath, process_folder)
                return True, process_folder

            if tarfile.is_tarfile(filepath):
                self.extract_tar(filepath, process_folder)
                return True, process_folder

            if is_rarfile(filepath):
                self.extract_rar(filepath, process_folder)
                return True, process_folder

            if self.is_sevenZfile(filepath):
                self.extract_sevenz(filepath, process_folder)
                return True, process_folder

        except Exception as exc:
            self.db_process.update_remark(unique_id, f"Extraction failed: {exc}")

        return False, None

    # --------------------------------------------------
    # MANUSCRIPT FORMAT (UNCHANGED)
    # --------------------------------------------------

    def check_manuscript_format(self, customer, file_list, unique_id, error_folder):
        if re.match("^SAGE$", customer):
            doc_found = bool(file_list.get("doc"))
            tex_found = bool(file_list.get("tex"))
            pdf_found = bool(file_list.get("pdf"))
            office_found = bool(file_list.get("office"))

            do_next = doc_found or tex_found or pdf_found or office_found
            return doc_found, do_next

        return False, False

    # --------------------------------------------------
    # FILE DETAILS (UNCHANGED LOGIC)
    # --------------------------------------------------

    def get_file_details(self, customer, file_list, error_folder, unique_id):
        journal_id = article_id = ms_no = None
        details_found = False

        process_folder = file_list.get("process_folder")
        if not process_folder:
            return False, None, None, None

        folder_name = os.path.basename(process_folder).split(".")[0]
        error_folder = os.path.join(error_folder, folder_name)

        if file_list.get("short_metadata"):
            for meta in file_list["short_metadata"]:
                try:
                    root = ET.parse(meta).getroot()
                    jid = root.find(
                        './front/journal-meta/journal-id/[@journal-id-type="acronym"]'
                    )
                    if jid is None:
                        jid = root.find(
                            './front/journal-meta/journal-id/[@journal-id-type="publisher"]'
                        )

                    journal_id = jid.text.upper()
                    article_id = root.find(
                        './front/article-meta/article-id[@pub-id-type="publisher-id"]'
                    ).text
                    ms_no = root.find(
                        './front/article-meta/article-id[@pub-id-type="manuscript"]'
                    ).text

                    info = GetArticleId()
                    ok, article_id, jrn_tla = info.smart_login(
                        article_id, ms_no, journal_id, process_folder
                    )
                    if ok:
                        journal_id = jrn_tla or journal_id
                        details_found = True
                        break
                except Exception:
                    pass

        if not details_found and file_list.get("long_metadata"):
            for meta in file_list["long_metadata"]:
                try:
                    root = ET.parse(meta).getroot()
                    journal_id = root.find(".//journal_abbreviation").text.upper()
                    ms_no = root.find("./article").attrib.get("ms_no")

                    info = GetArticleId()
                    ok, article_id, jrn_tla = info.smart_login(None, ms_no, journal_id, process_folder)
                    if ok:
                        journal_id = jrn_tla or journal_id
                        details_found = True
                        break
                except Exception:
                    pass

        if not details_found:
            if os.path.exists(error_folder):
                shutil.rmtree(error_folder, ignore_errors=True)
            shutil.move(process_folder, error_folder)
            self.db_process.update_remark(
                unique_id, "Metadata file not found in package"
            )
            return False, None, None, None

        self.db_process.update_data(
            customer, ms_no, journal_id, article_id, unique_id
        )
        return True, journal_id, article_id, ms_no

    def create_info_file(self, customer, file_dics, jid, aid, ms_no, unique_id):
        if re.match("^SAGE$", customer):
            processFolder = file_dics['process_folder']
            file_info = processFolder + "/MergeInfo.xml"

            if 'long_metadata' in file_dics and file_dics['long_metadata']:
                if len(file_dics['long_metadata']) > 0:
                    long_meta = file_dics['long_metadata'][0]
                    with open(long_meta, "r", encoding="utf8") as file:
                        file_text = file.read()
                        if bool(re.search("<article_set", file_text)) is False:
                            file_dics['long_metadata'] = None
                else:
                    file_dics['long_metadata'] = None
            else:
                # long_metadata is missing or empty list — treat as None
                file_dics['long_metadata'] = None

            if "short_metadata" in file_dics.keys():
                short_metadata = file_dics['short_metadata']
                if len(short_metadata) > 0:
                    short_metadata = file_dics['short_metadata'][0]
                else:
                    file_dics['short_metadata'] = None

            if file_dics['long_metadata'] is None:
                merger_input = self.analyser_config['FOLDERS']['MERGER']
                merger_input = re.sub(r"\[CUSTOMER\]", customer, merger_input)
                merger_input = re.sub(r"\[JID\]", jid, merger_input)
                merger_input = re.sub(r"\[AID\]", aid, merger_input)

                uid_file = os.path.join(merger_input, "unique_id.json")
                json_file_name = jid + "_" + aid + ".json"
                json_process_path = os.path.join(processFolder, json_file_name)
                json_merger_path = os.path.join(merger_input, json_file_name)

                if not os.path.exists(merger_input):
                    pathlib.Path(merger_input).mkdir(parents=True, exist_ok=True)

                with open(uid_file, "w") as uid:
                    uuid = "\"unique_id:\" \"" + unique_id + "\""
                    uid.write(uuid)

                shutil.copy(json_process_path, json_merger_path)

                doc_files = file_dics['doc']
                doc_dic = {}
                normalizer_input = {}

                for doc in doc_files:
                    doc_dic[os.path.split(doc)[1]] = doc

                # ── Short-metadata only + single docx: skip merger ────────
                # Rename the file to JID_AID_CLN.docx and place directly
                # into ParaStyler input folder. No merge needed.
                has_short_meta_only = (
                    file_dics.get('short_metadata') is not None
                    and len(doc_dic) == 1
                )

                if has_short_meta_only:
                    # Derive ParaStyler input path from merger path
                    # Merger:     .../Merger_INPUT/[CUSTOMER]/JID_AID/
                    # ParaStyler: .../ParaStyler_INPUT/[CUSTOMER]/JID_AID/
                    parastyler_input = re.sub(
                        r"Merger_INPUT", "ParaStyler_INPUT",
                        merger_input, flags=re.IGNORECASE
                    )
                    if not os.path.exists(parastyler_input):
                        pathlib.Path(parastyler_input).mkdir(parents=True, exist_ok=True)

                    single_docx_name = list(doc_dic.keys())[0]
                    single_docx_path = doc_dic[single_docx_name]
                    cln_name = jid + "_" + aid + "_CLN.docx"
                    dest_file = os.path.join(parastyler_input, cln_name)
                    shutil.copy(single_docx_path, dest_file)

                    # Also copy JSON to ParaStyler input
                    json_parastyler_path = os.path.join(parastyler_input, json_file_name)
                    shutil.copy(json_process_path, json_parastyler_path)

                    # Write unique_id to ParaStyler folder
                    ps_uid_file = os.path.join(parastyler_input, "unique_id.json")
                    with open(ps_uid_file, "w") as uid:
                        uuid = "\"unique_id:\" \"" + unique_id + "\""
                        uid.write(uuid)

                    normalizer_input['selected'] = {cln_name: dest_file}
                    normalizer_input['customer'] = customer
                    normalizer_input['folder'] = os.path.split(parastyler_input)[1]
                    normalizer_input['process'] = "mParaStyler"
                    normalizer_input['skip_merger'] = True
                    process_auto = True

                    print(f"[INFO] Short-metadata only package with single docx. "
                          f"Renamed '{single_docx_name}' -> '{cln_name}', "
                          f"placed in ParaStyler input (skipping merger).")
                    return process_auto, None, normalizer_input

                if len(doc_files) == 0:
                    normalizer_input['selected'] = {}
                    for docf in doc_dic:
                        merger_file = os.path.join(merger_input, docf)
                        source_path = doc_dic[docf]
                        shutil.copy(source_path, merger_file)
                        normalizer_input['selected'][docf] = merger_file
                        merger_fldr = os.path.split(merger_input)[1]
                        normalizer_input['customer'] = customer
                        normalizer_input['folder'] = merger_fldr
                        normalizer_input['process'] = "mMerge"
                        process_auto = False
                else:
                    for doc_file in doc_dic:
                        for_merging = doc_dic[doc_file]
                        merge_file = os.path.split(for_merging)[1]
                        dest_file = os.path.join(merger_input, merge_file)
                        shutil.copy(for_merging, dest_file)
                        if 'selected' in normalizer_input:
                            normalizer_input['selected'][merge_file] = dest_file
                        else:
                            normalizer_input['selected'] = {}
                            normalizer_input['selected'][merge_file] = dest_file

                    merger_fldr = os.path.split(merger_input)[1]
                    normalizer_input['customer'] = customer
                    normalizer_input['folder'] = merger_fldr
                    normalizer_input['process'] = "mMerge"

                    if len(doc_dic) == 1:
                        process_auto = True
                    else:
                        process_auto = False

                return process_auto, None, normalizer_input

            elif "long_metadata" in file_dics.keys() and file_dics['long_metadata']:
                if len(file_dics['long_metadata']) > 0:
                    long_metadata = file_dics['long_metadata'][0]

                with open(long_metadata, "r", encoding="utf8") as file:
                    file_text = file.read()
                    file_text = re.sub(r"<\!--s1(.+?)-->", "", file_text)
                    file_text = re.sub(r"<\!DOCTYPE(.+?)>", "", file_text)

                with open(long_metadata, "w", encoding="utf8") as file:
                    file.write(file_text)

                xmltrns = XmlTransform()
                xslpath = os.path.dirname(os.path.realpath(__file__)) + "/xsl/metaToFileInfo.xsl"
                proceed, info_xml = xmltrns.transform_xml(long_metadata, file_info, xslpath)

                if proceed is True:
                    with open(info_xml, encoding="utf8") as xml_file:
                        data_dict = xmltodict.parse(xml_file.read())
                        json_data = json.loads(json.dumps(data_dict))

                    file_list = json_data['root']['ItemType']
                    true_list, false_list, none_list = {}, {}, {}
                    doc_types = self.analyser_config['DocTypes']

                    if isinstance(file_list, dict):
                        file_list = [file_list]

                    for file in file_list:
                        file_merge = file['ItemDescription']['@merge']
                        file_type = file['ItemDescription']['@order_type']
                        file_name = file['FileName']

                        for doc_type in doc_types:
                            if file_type == doc_type:
                                if re.search("true", file_merge, re.IGNORECASE):
                                    true_list.setdefault(doc_type, []).append(file_name)
                                elif re.search("false", file_merge, re.IGNORECASE):
                                    false_list.setdefault(doc_type, []).append(file_name)
                                elif re.search("none", file_merge, re.IGNORECASE):
                                    none_list.setdefault(doc_type, []).append(file_name)

                    merger_input = self.analyser_config['FOLDERS']['MERGER']
                    merger_input = re.sub(r"\[CUSTOMER\]", customer, merger_input)
                    merger_input = re.sub(r"\[JID\]", jid, merger_input)
                    merger_input = re.sub(r"\[AID\]", aid, merger_input)

                    uid_file = os.path.join(merger_input, "unique_id.json")
                    info_file_name = os.path.split(info_xml)[1]
                    json_file_name = jid + "_" + aid + ".json"
                    json_process_path = os.path.join(processFolder, json_file_name)
                    json_merger_path = os.path.join(merger_input, json_file_name)
                    info_merger_path = os.path.join(merger_input, info_file_name)

                    if not os.path.exists(merger_input):
                        pathlib.Path(merger_input).mkdir(parents=True, exist_ok=True)

                    with open(uid_file, "w") as uid:
                        uuid = "\"unique_id:\" \"" + unique_id + "\""
                        uid.write(uuid)

                    shutil.copy(info_xml, info_merger_path)
                    shutil.copy(json_process_path, json_merger_path)

                    doc_dic = {os.path.split(doc)[1]: doc for doc in file_dics['doc']}
                    normalizer_input = {}
                    meta_ok = True

                    merger_order = self.analyser_config['MergeOrder']
                    for key, value in sorted(merger_order.items()):
                        if value in true_list:
                            for name in true_list[value]:
                                if name not in doc_dic:
                                    meta_ok = False
                                    break

                    # If only one doc/docx file exists (cumulative count),
                    # auto-merge regardless of merge flags — single file
                    # doesn't need manual selection
                    single_doc = (len(doc_dic) == 1)

                    if single_doc:
                        normalizer_input['selected'] = {}
                        for docf in doc_dic:
                            dest = os.path.join(merger_input, docf)
                            shutil.copy(doc_dic[docf], dest)
                            normalizer_input['selected'][docf] = dest
                        process_auto = True
                    elif false_list or not true_list or meta_ok is False:
                        normalizer_input['selected'] = {}
                        for docf in doc_dic:
                            dest = os.path.join(merger_input, docf)
                            shutil.copy(doc_dic[docf], dest)
                            normalizer_input['selected'][docf] = dest
                        process_auto = False
                    else:
                        normalizer_input['selected'] = {}
                        for key, value in sorted(merger_order.items()):
                            if value in true_list:
                                for name in true_list[value]:
                                    dest = os.path.join(merger_input, name)
                                    shutil.copy(doc_dic[name], dest)
                                    normalizer_input['selected'][name] = dest
                        process_auto = True

                    normalizer_input['customer'] = customer
                    normalizer_input['folder'] = os.path.split(merger_input)[1]
                    normalizer_input['process'] = "mMerge"

                    return process_auto, info_xml, normalizer_input

                else:
                    # XSLT transform failed — fall back to no-metadata path:
                    # copy doc files to merger input and let the user merge manually.
                    print(f"[WARN] XSLT transform failed for {long_metadata}, "
                          f"falling back to manual merge.")
                    merger_input = self.analyser_config['FOLDERS']['MERGER']
                    merger_input = re.sub(r"\[CUSTOMER\]", customer, merger_input)
                    merger_input = re.sub(r"\[JID\]", jid, merger_input)
                    merger_input = re.sub(r"\[AID\]", aid, merger_input)

                    if not os.path.exists(merger_input):
                        pathlib.Path(merger_input).mkdir(parents=True, exist_ok=True)

                    uid_file = os.path.join(merger_input, "unique_id.json")
                    with open(uid_file, "w") as uid:
                        uuid = "\"unique_id:\" \"" + unique_id + "\""
                        uid.write(uuid)

                    json_file_name = jid + "_" + aid + ".json"
                    json_process_path = os.path.join(processFolder, json_file_name)
                    json_merger_path = os.path.join(merger_input, json_file_name)
                    shutil.copy(json_process_path, json_merger_path)

                    doc_dic = {os.path.split(doc)[1]: doc for doc in file_dics['doc']}
                    normalizer_input = {}
                    normalizer_input['selected'] = {}
                    for docf in doc_dic:
                        dest = os.path.join(merger_input, docf)
                        shutil.copy(doc_dic[docf], dest)
                        normalizer_input['selected'][docf] = dest
                    normalizer_input['customer'] = customer
                    normalizer_input['folder'] = os.path.split(merger_input)[1]
                    normalizer_input['process'] = "mMerge"

                    return False, None, normalizer_input

    def run_analyser(self, filepath, customer, errorFolder, unique_id):
        isAuto = False
        infFile = None
        nInput = None
        word_file_found = False
        details_found = False

        ok, process_path = self.move_to_process(filepath, unique_id)
        time.sleep(10)

        if ok is True:
            file_list = self.get_file_list(process_path)

            details_found, jid, aid, ms_no = self.get_file_details(
                customer, file_list, errorFolder, unique_id
            )
            if details_found is True:
                word_file_found, do_next = self.check_manuscript_format(
                    customer, file_list, unique_id, errorFolder
                )
                if word_file_found is True:
                    result = self.create_info_file(
                        customer, file_list, jid, aid, ms_no, unique_id
                    )
                    if result is not None:
                        isAuto, infFile, nInput = result
                    else:
                        print(f"[ERROR] create_info_file returned None for {jid}_{aid}")

                    self.db_process.update_db(
                        unique_id, "mAnalyzer", "COMPLETED", "", "No error"
                    )

                else:
                    merger_input = self.analyser_config["FOLDERS"]["MERGER"]
                    merger_input = re.sub(r"\[CUSTOMER\]", customer, merger_input)
                    merger_input = re.sub(r"\[JID\]", jid, merger_input)
                    merger_input = re.sub(r"\[AID\]", aid, merger_input)

                    try:
                        shutil.move(process_path, merger_input)
                    except Exception as e:
                        print(e)

        if process_path and os.path.exists(process_path):
            try:
                shutil.rmtree(process_path, ignore_errors=True)
            except Exception:
                pass
        return details_found, isAuto, infFile, nInput, word_file_found



# analyser = mAnalyzer()
# analyser.run_analyser('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-02-15-22-03-27.zip', "SAGE", "V:\\FOR_BREAKDOWN\\ERROR", "sfjskfusikk999kfsdf")
# analyser.get_file_details('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-02-15-22-03-27.zip', "SAGE", "V:\\FOR_BREAKDOWN\\ERROR", "sfjskfusikk999kfsdf")
# analyser.create_info_file('SAGE', {'process_folder': 'V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-02-15-22-03-27', 'short_metadata': [], 'long_metadata': ['V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-02-15-22-03-27/tab-25-09-199-metadata.xml'], 'manifest': ['V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-02-15-22-03-27/tab-25-09-199-manifest.html'], 'doc': ['V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-02-15-22-03-27/TAB-25-09-199R1-Manuscript_Clean.docx'], 'office': [], 'graphics': [], 'tex': [], 'pdf': [], 'merge_info': [], 'unknown': []}, 'TAB', '1412846', 'TAB-25-09-199.R1', '7f985fde1e8d4707853c71147922f7d6')
# file_dic = {'process_folder': 'V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-01-27-21-09-26', 'short_metadata': [], 'long_metadata': ['V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-01-27-21-09-26/shm-25-0455-metadata.xml'], 'manifest': ['V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-01-27-21-09-26/shm-25-0455-manifest.html'], 'doc': ['V:/FOR_BREAKDOWN/PROCESS/Article_Attachments-2026-01-27-21-09-26/SHM-25-0455-Revised Manuscript.docx'], 'office': [], 'graphics': [], 'tex': [], 'pdf': [], 'merge_info': [], 'unknown': []}
# isAuto, infFile, nInput = analyser.create_info_file("SAGE", file_dic, "SHM", "1381259", "SHM-25-0455.R1", "sfjskfusikk999kfsdf")
# analyser.run_analyser('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-01-27-21-09-26.zip', "SAGE", "V:\\FOR_BREAKDOWN\\ERROR", "sfjskfusikk999kfsdf")
# file_list = analyser.get_file_list('V:\\FOR_BREAKDOWN\\PROCESS\\Article_Attachments-2026-01-26-14-18-17')
# details_found, jid, aid, ms_no = analyser.get_file_details("SAGE", file_list, "ERROR", "ABC")
# analyser.create_info_file('SAGE', file_list, 'HPI', '1416040', 'HIPINT-25-0181.R2', '7f985fde1e8d4707853c71147922f7d6')
# details_found, jid, aid, ms_no = analyser.get_file_details("SAGE", file_list, "ERROR", "ABC")
# print(file_list)
# analyser.move_to_process('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-01-26-14-18-17.zip', "sfjskfusikk999kfsdf")
# analyser.extract_rar('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-01-26-14-18-17.zip', '\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-01-26-14-18-17')
# analyser.run_analyser('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\Article_Attachments-2026-01-26-14-18-17.zip', "SAGE", "V:\\FOR_BREAKDOWN\\ERROR", "sfjskfusikk999kfsdf")
# filelist = analyser.get_file_list('V:\\FOR_BREAKDOWN\\PROCESS\\Attachment_TYPE-Article_2023')
# # proceed, info_xml = xmltrns.transform_xml(long_metadata, file_info, "./xsl/metaToFileInfo.xsl")
# from TransformXml import XmlTransform
# xmltrns = XmlTransform()
# long_metadata = 'V:\\FOR_BREAKDOWN\\MERGER_INPUT\\SAGE\\SGO_1264348\\so-24-0117-metadata.xml'
# long_metadata = xmltrns.strip_dtd_declaration(long_metadata)
# xslpath = os.path.dirname(os.path.realpath(__file__)) + "/xsl/metaToFileInfo.xsl"
# file_info = 'V:\FOR_BREAKDOWN\MERGER_INPUT\SAGE\SGO_1264348\MergeInfo.xml'
# if os.path.exists(long_metadata):
#     proceed, info_xml = xmltrns.transform_xml(long_metadata, file_info, xslpath)
