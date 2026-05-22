import os
import re
import logging
import subprocess
from loadconfig import getconfig
from pyLabelAuthor import PyLabelAuthor
from charConverter import HexEscapeConverter

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class LabelAuthor:
    def __init__(self):
        self.file_type = ''
        self.customer = ''
        self.configFolder, self.breakDownConfig = getconfig()
        self.author_process_exe = os.path.join(self.configFolder, 'SupportingFiles', 'au_tag.exe')
        # self.author_process_exe = os.path.join(self.configFolder, 'SupportingFiles', 'au_tag.exe')
        self.degree_xml = os.path.join(self.configFolder, 'SupportingFiles', 'degree.xml')
        self.converter = HexEscapeConverter()

    def author_process(self, author):
        converted_author = self.converter.convert_non_ascii_to_hex_escapes(author)
        input_args = [self.degree_xml, converted_author]
        command = [self.author_process_exe] + input_args
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True)
            gauthor = result.stdout
            gauthor = self.converter.convert_hex_escapes_to_utf8(gauthor)
            # au_error = result.stderr
            # return_code = result.returncode
        except subprocess.CalledProcessError as e:
            py_label = PyLabelAuthor()
            ret_aut, only_aut = py_label.author_process(author)
            return ret_aut, only_aut
        except FileNotFoundError:
            logger.error("Executable not found: %s", self.author_process_exe)
            return None

        # Process output
        pyLabel = PyLabelAuthor()
        gauthor, only_authors = pyLabel.author_post_process(gauthor)
        au_count = 0
        pattern = r"<au>((?:[^<]*|<(?!/au))+)</au>"
        while re.search(pattern, gauthor, re.I):
            au_count += 1
            gauthor = re.sub(pattern, rf'\1<span style="color: blue;">[AU{au_count}]</span>', gauthor, 1, re.IGNORECASE)

        aut_dict = {}
        au_count = 0
        pattern = re.compile(r"<au>(?P<name>((?:[^<]*|<(?!/au))+))</au>")
        matches = pattern.findall(only_authors)
        for match in matches:
            name, full_name = match
            au_count += 1
            aut_dict[au_count] = name

        # display_label = gauthor
        # display_label = re.sub('<span style="color: blue;">', '', display_label)
        # display_label = re.sub('</span>', '\n', display_label)
        # logger.info("Authors:\n%s\nAny issues in labeling please check degrees and separator", display_label)
        return gauthor, aut_dict


# Example usage
# proc_auth = LabelAuthor()
# aut, only_aut = proc_auth.author_process("Lingli Qiu<sup>1,2</sup>, Wenqiang Zhang<sup>1</sup>, Zhixin Tan<sup>1</sup>, Xuan Wu<sup>1</sup>, Yutong Wang<sup>1</sup>, Mingshuang Tang<sup>1</sup>, Lin Chen<sup>1</sup>, Yanqiu Zou<sup>1</sup>, Yunjie Liu<sup>1</sup>, Bowen Lei<sup>1</sup>, Xiaofeng Ma<sup>1</sup>, Di Zhang<sup>1</sup>, Wenzhi Wang<sup>3</sup>, Yiping Jia<sup>4</sup>, Qiurong He<sup>5</sup>, Lei Sun<sup>3</sup>, Lu Wang<sup>3</sup>, Jian Xu<sup>3</sup>, Yao Chen<sup>3</sup>, Mengyu Fan<sup>1</sup>, Jiayuan Li<sup>1</sup>, Ben Zhang<sup>1,6</sup>, Xiaoling Wen<sup>7*</sup>, and Xia Jiang<sup>1,8,9*</sup>")
# print(aut)
# print(only_aut)
