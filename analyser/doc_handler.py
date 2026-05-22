import os
import pathlib
import re

class DocHandler:
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
            "unknown": []
        }

        for path, _, files in os.walk(process_folder):
            re_doc = r"\\doc|/doc"
            re_graphics = r"\\(graphics|images)|/(graphics|images)"
            
            for name in files:
                if name.startswith("~$"):
                    continue
                    
                filename = os.path.join(path, name)
                file_ext = pathlib.Path(name).suffix.lower()
                gra_path = bool(re.search(re_graphics, path, re.IGNORECASE))
                gra_ext = file_ext in ['.tif', '.tiff', '.jpg', '.jpeg', '.eps', '.svg', '.png']

                # Categorize files
                if re.match(r"SAGE-metadata-([^.]+).xml", name) or re.match(r"([0-9]{2})-([0-9]+).xml", name):
                    input_files["short_metadata"].append(filename)
                elif re.match(r"(.+?)-metadata.xml", name):
                    input_files["long_metadata"].append(filename)
                elif re.match(r"(.+?)-manifest.html", name):
                    input_files["manifest"].append(filename)
                elif file_ext in ['.docx', '.doc']:
                    input_files["doc"].append(filename)
                elif file_ext in ['.xls', '.xlsx', '.ppt', '.pptx']:
                    input_files["office"].append(filename)
                elif gra_path or gra_ext:
                    input_files["graphics"].append(filename)
                elif re.search(re_doc, path, re.IGNORECASE) or file_ext == '.tex':
                    input_files["tex"].append(filename)
                elif file_ext == '.pdf':
                    input_files["pdf"].append(filename)
                elif "MergeInof" in name:
                    input_files["merge_info"].append(filename)
                else:
                    input_files["unknown"].append(filename)
                    
        return input_files

    def check_manuscript_format(self, customer, file_list):
        if customer != "SAGE":
            return False, False

        doc_files_found = bool(file_list.get("doc"))
        tex_files_found = bool(file_list.get("tex"))
        pdf_files_found = bool(file_list.get("pdf"))
        office_files_found = bool(file_list.get("office"))
        do_next = False

        if not doc_files_found:
            error_messages = {
                "tex": "[0014]: Only Tex files found - manual process needed",
                "pdf": "[0014]: Only PDF files found - manual process needed",
                "office": "[0014]: Only office files found - manual process needed"
            }
            
            for file_type, message in error_messages.items():
                if locals()[f"{file_type}_files_found"]:
                    print(message)
                    break
            else:
                print("[0014]: Check manuscript files - manual process needed")
        else:
            self._handle_doc_files(file_list)

        if any([doc_files_found, tex_files_found, pdf_files_found, office_files_found]):
            do_next = True
            
        return doc_files_found, do_next

    def _handle_doc_files(self, file_list):
        word_documents = [
            f for f in file_list.get("doc", [])
            if pathlib.Path(f).suffix.lower() in ['.docx', '.doc']
        ]
        return bool(word_documents)