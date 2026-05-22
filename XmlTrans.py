import lxml.etree as ET


class TransformXml:
    def xml_trans(self, xml_filename, xsl_filename, output_filename):
        dom = ET.parse(xml_filename)
        xslt = ET.parse(xsl_filename)
        transform = ET.XSLT(xslt)
        newdom = transform(dom)
        out_xml = ET.tostring(newdom, pretty_print=True)
        with open(output_filename, "w") as f:
            f.write(out_xml)