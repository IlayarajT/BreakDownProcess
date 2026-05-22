import os
import re
import shutil
import subprocess
import zipfile


class XmlTransformJar:
    def __init__(self):
        pass

    def transform_xml(self, input_file, output_file, xsl_file):
        try:
            command = [
                'java', '-jar', "ParaStyler/saxon9pe.jar",
                '-s:' + input_file,  # Source XML file
                '-xsl:' + xsl_file,  # XSLT file
                '-o:' + output_file  # Output XML file
            ]
            result = subprocess.run(command, capture_output=True, text=True)
            if result.returncode != 0:
                print(result.stderr)
                return False, 'Transformation failed'
            return True, output_file
        except Exception as e:
            print(e)
            return False, 'No file info'

    def udpate_table_cells(self, docxfile):
        print(f"update table cells {docxfile}")
        with zipfile.ZipFile(docxfile, 'r') as zip_ref:
            zip_ref.extract('word/document.xml', 'temp')

        if not os.path.exists("D:/mProjects/test"):
            os.mkdir("D:/mProjects/test")
        if os.path.exists('D:/mProjects/test/document.xml'):
            os.remove('D:/mProjects/test/document.xml')

        result_file = 'D:/mProjects/test/document.xml'
        try:
            command = [
                'java', '-jar', "ParaStyler/saxon9pe.jar",
                '-s:temp/word/document.xml',  # Source XML file
                '-xsl:xsl/tableFormat.xsl',  # XSLT file
                '-o:' + result_file  # Output XML file
            ]
            result = subprocess.run(command, capture_output=True, text=True)
            if result.returncode != 0:
                print(result.stderr)
                return
            os.remove('temp/word/document.xml')
            shutil.copy(result_file, "temp/word/document.xml")
            with zipfile.ZipFile(docxfile, 'r') as zip_read:
                with zipfile.ZipFile(docxfile + '.tmp', 'w') as zip_write:
                    for item in zip_read.infolist():
                        if item.filename != 'word/document.xml':
                            buffer = zip_read.read(item.filename)
                            zip_write.writestr(item, buffer)
            with zipfile.ZipFile(docxfile + '.tmp', 'a') as zip_ref:
                zip_ref.write('temp/word/document.xml', 'word/document.xml')
            os.remove('temp/word/document.xml')
            os.replace(docxfile + '.tmp', docxfile)
        except Exception as e:
            print(e)

    def trans_xml(self, input_file, xslt_file, output_file):
        try:
            command = [
                'java', '-jar', "ParaStyler/saxon9pe.jar",
                '-s:' + input_file,  # Source XML file
                '-xsl:' + xslt_file,  # XSLT file
                '-o:' + output_file  # Output XML file
            ]
            result = subprocess.run(command, capture_output=True, text=True)
            if result.returncode != 0:
                print(result.stderr)
                return None
            return output_file
        except Exception as e:
            print(e)
            return None

    def jar_transform(self, input_file, xslt_file, output_file):
        try:
            command = [
                'java', '-jar', "ParaStyler/saxon9pe.jar",
                '-s:' + input_file,  # Source XML file
                '-xsl:' + xslt_file,  # XSLT file
                '-o:' + output_file  # Output HTML file
            ]
            result = subprocess.run(command, capture_output=True, text=True)
            return result
        except Exception as e:
            print(e)
            return None

    def strip_dtd_declaration(self, input_file):
        with open(input_file, 'r', encoding='utf-8') as file:
            content = file.read()
        content = re.sub(r'<!DOCTYPE[^>]*>', '', content)
        stripped_file = re.sub(r'\.xml$', '_stripped.xml', input_file, flags=re.IGNORECASE)
        with open(stripped_file, 'w', encoding='utf-8') as file:
            file.write(content)
        return stripped_file

# trns_xml = XmlTransformJar()
# trns_xml.trans_xml("V:\FOR_BREAKDOWN\BreakDown_INPUT\SAGE\SGO_1264348\docs\SGO_1264348_CLN_AS\word\document.xml",
#                    "xsl/postAs.xsl",
#                    "V:\FOR_BREAKDOWN\BreakDown_INPUT\SAGE\SGO_1264348\docs\SGO_1264348_CLN_AS\word\document1.xml")
# trns_xml.trans_xml("V:\FOR_BREAKDOWN\BreakDown_INPUT\SAGE\SGO_1264348\\author_line.htm",
#                    "xsl/htmlToDocxXml.xsl",
#                    "V:\FOR_BREAKDOWN\BreakDown_INPUT\SAGE\SGO_1264348\\author_line.xml")
# trns_xml.jar_transform("temp/paragraph_1.xml", "xsl/docxXmlToHtml.xsl", "temp/output_1.htm")
