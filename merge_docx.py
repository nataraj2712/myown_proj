# coding: utf-8

import base64
import collections
import io
import itertools
import os,re
import zipfile
import sys
from six.moves.urllib import request as rq
# import six.moves.urllib.request
from lxml import etree

PackagePart = collections.namedtuple('PackagePart', ['uri', 'content_type', 'data'])


class ContentTypes(object):
    NS = {'ct': u"http://schemas.openxmlformats.org/package/2006/content-types"}

    def __init__(self):
        self._defaults = {}
        self._overrides = {}

    def parse_xml_data(self, data):
        tree = etree.fromstring(data)  # type: etree._Element
        self._defaults = {n.attrib[u'Extension']: n.attrib[u'ContentType']
                          for n in tree.xpath(u'//ct:Default', namespaces=self.NS)}
        self._overrides = {n.attrib[u'PartName']: n.attrib[u'ContentType']
                           for n in tree.xpath(u'//ct:Override', namespaces=self.NS)}

    def resolve(self, part_name):
        basename = os.path.basename(part_name)
        ext = basename.rsplit(".", 1)[1]
        content_type = self._overrides.get(part_name) or self._defaults.get(ext)
        return content_type


def iter_package(zip_path):
    content_types = ContentTypes()
    with zipfile.ZipFile(zip_path, mode="r") as f:
        for name in f.namelist():
            if name == "[Content_Types].xml":
                content_types.parse_xml_data(f.read(name))
            else:
                uri = "/" + rq.pathname2url(name)
                content_type = content_types.resolve(uri)
                data = f.read(name)
                yield PackagePart(uri, content_type, data)


def opc_to_flat_opc(src_docx_path, dst_opc_path):
    pkg = u"http://schemas.microsoft.com/office/2006/xmlPackage"

    ext = os.path.splitext(src_docx_path)[1].lower()
    progid = {'.docx': u"Word.Document",
              '.xlsx': u"Excel.Sheet",
              '.pptx': u"PowerPoint.Show"}[ext]

    content = (u'<?mso-application progid="{progid}"?>'
               u'<pkg:package xmlns:pkg="{pkg}"/>').format(progid=progid, pkg=pkg)

    document = etree.parse(io.StringIO(content))  # type: etree._ElementTree
    root = document.getroot()

    ns = {'pkg': pkg}

    for part in iter_package(src_docx_path):
        node = etree.SubElement(root, u"{{{pkg}}}part".format(pkg=pkg), nsmap=ns)
        node.attrib[u"{{{pkg}}}name".format(pkg=pkg)] = part.uri
        node.attrib[u"{{{pkg}}}contentType".format(pkg=pkg)] = part.content_type
        if part.content_type.endswith("xml"):
            data = etree.SubElement(node, u"{{{pkg}}}xmlData".format(pkg=pkg), nsmap=ns)
            data.append(etree.fromstring(part.data))
        else:
            node.attrib[u"{{{pkg}}}compression".format(pkg=pkg)] = "store"
            data = etree.SubElement(node, u"{{{pkg}}}binaryData".format(pkg=pkg), nsmap=ns)
            encoded = base64.b64encode(part.data).decode()  # bytes -> str
            iterable = iter(encoded)
            chunks = list(iter(lambda: list(itertools.islice(iterable, 76)), []))
            chunks = u"\n".join(u"".join(chunk) for chunk in chunks)
            data.text = chunks

    content = etree.tostring(document,
                             xml_declaration=True,
                             encoding='UTF-8',
                             pretty_print=False,
                             with_tail=False,
                             standalone=True)
    with io.open(dst_opc_path, mode="wb") as f:
        f.write(content)


inFile = sys.argv[1]
outFile = re.sub('.docx\Z','.xml',inFile,flags=re.I)
opc_to_flat_opc(inFile,outFile)
