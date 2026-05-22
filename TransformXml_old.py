#from saxonpy import *
# import os
# from nodekind import *
#import sys
#sys.path.append("C:/Program Files/Saxonica/SaxonPEC 11.4/Saxon.C.API/python-saxon")
import os
import re
import shutil
import subprocess
import zipfile
import atexit

from saxonc import PySaxonProcessor

# ---------------------------------------------------------------------------
# Module-level Saxon processor — ONE JVM for the entire process lifetime.
#
# Rules enforced here:
#   1. PySaxonProcessor() created ONCE — JNI_CreateJavaVM can only be called
#      once per process (second call → error -5 JVM_EEXIST).
#   2. __enter__ called once to initialise the JET runtime.
#   3. __exit__ registered with atexit → DetachCurrentThread fires on clean
#      process exit (prevents jet_err "Thread terminated without notifying JVM").
#   4. A NEW Xslt30Processor is created for every transform call —
#      reusing one across calls causes EXCEPTION_ACCESS_VIOLATION (0xC0000005)
#      because the internal JET object pointers go stale after transform_to_file.
#   5. proc.exception_clear() is called after every transform to flush any
#      residual JET error state on the shared processor object.
# ---------------------------------------------------------------------------
_saxon_proc = PySaxonProcessor(license=False)
_saxon_proc.__enter__()
atexit.register(_saxon_proc.__exit__, None, None, None)


def _get_proc() -> PySaxonProcessor:
    """Return the shared Saxon processor instance."""
    return _saxon_proc


def _run_xslt(stylesheet_file: str,
              source_file: str,
              output_file: str,
              parameters: dict = None) -> None:
    """
    Execute a single XSLT 3.0 transform.
    Always creates a fresh Xslt30Processor and clears processor state
    afterwards — never reuses an old processor object across calls.
    """
    proc = _get_proc()
    xslt = proc.new_xslt30_processor()          # fresh object every time
    try:
        if parameters:
            for name, value in parameters.items():
                xslt.set_parameter(name, value)
        xslt.transform_to_file(
            stylesheet_file=stylesheet_file,
            source_file=source_file,
            output_file=output_file,
        )
    finally:
        proc.exception_clear()                   # flush stale JET state


class XmlTransform:
    def __init__(self):
        pass

    def transform_xml(self, input_file, output_file, xsl_file):
        print(_get_proc().version)
        if os.path.exists(input_file):
            print(input_file)
        try:
            _run_xslt(
                stylesheet_file=xsl_file,
                source_file=input_file,
                output_file=output_file,
            )
            return True, output_file
        except Exception as e:
            print(e)
            return False, 'No file info'

    def udpate_table_cells(self, docxfile):
        print(f"update table cells f{docxfile}")

        with zipfile.ZipFile(docxfile, 'r') as zip_ref:
            zip_ref.extract('word/document.xml', 'temp')

        result_file = 'D:/mProjects/test/document.xml'
        if not os.path.exists("D:/mProjects/test"):
            os.mkdir("D:/mProjects/test")
        if os.path.exists(result_file):
            os.remove(result_file)

        # Use Java Saxon jar — saxonc native DLL crashes on this machine
        result = self.jar_transform(
            input_file='temp/word/document.xml',
            xslt_file='xsl/tableFormat.xsl',
            output_file=result_file,
        )
        if result.returncode != 0:
            print(f"[jar_transform error] {result.stderr}")
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

    def trans_xml(self, input_file, xslt_file, output_file):
        _run_xslt(
            stylesheet_file=xslt_file,
            source_file=input_file,
            output_file=output_file,
        )
        return output_file

    def jar_transform(self, input_file, xslt_file, output_file):
        command = [
            'java', '-jar', "ParaStyler/saxon9pe.jar",
            '-s:' + input_file,
            '-xsl:' + xslt_file,
            '-o:' + output_file,
        ]
        result = subprocess.run(command, capture_output=True, text=True)
        return result

    def jar_transform_xml(self, input_file, xslt_file, output_file):
        command = [
            'java', '-jar', "ParaStyler/saxon9pe.jar",
            '-x',
            '-s:' + input_file,
            '-xsl:' + xslt_file,
            '-o:' + output_file,
        ]
        result = subprocess.run(command, capture_output=True, text=True)
        return result

    def strip_dtd_declaration(self, input_file):
        with open(input_file, 'r', encoding='utf-8') as file:
            content = file.read()
        content = re.sub(r'<!DOCTYPE[^>]*>', '', content)
        stripped_file = re.sub(r'\.xml$', '_stripped.xml', input_file,
                                flags=re.IGNORECASE)
        with open(stripped_file, 'w', encoding='utf-8') as file:
            file.write(content)
        return stripped_file


# trns_xml = XmlTransform()
# trns_xml.jar_transform_xml("temp\\author_line.htm", "xsl/htmlToDocxXml.xsl", "temp\\author_line.htm")
