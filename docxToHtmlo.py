import argparse
import os

import mammoth


class DocxToHtml:
    def __init__(self):
        self.parser = argparse.ArgumentParser()
        self.parser.add_argument('-f', '--filepath', help="Provide file path to process", type=str)
        self.args = self.parser.parse_args()

    def html_convert(self):
        docx_file = self.args.filepath
        html_file = self.args.filepath
        pre, ext = os.path.splitext(html_file)
        html_file = pre + ".html"
        input_docx = open(docx_file, 'rb')
        output_html = open(html_file, 'wb')
        document = mammoth.convert_to_html(input_docx)
        output_html.write(document.value.encode('utf8'))
        print(html_file)
        input_docx.close()
        output_html.close()


html_conv = DocxToHtml()
html_conv.html_convert()

