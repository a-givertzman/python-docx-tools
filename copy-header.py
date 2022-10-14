import os
import re
from xml.dom import XML_NAMESPACE
from xml.etree.ElementTree import QName
import shutil
import zipfile
import pathlib
import lxml.etree

class DocXType:
    header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    xmlns = {
        'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }


class DocX:
    def __init__(
        self,
        path: str,
    ) -> None:
        self._docPath = path
        self._docDir = str(pathlib.Path(path).parent)
        self._docName = pathlib.Path(path).name
        self._tmpDir = f'{self._docDir}/tmp/{self._docName}'
        self._wordDir = f'{self._docDir}/tmp/{self._docName}/word'
        self._wordDocPath = f'{self._wordDir}/document.xml'
        self._wordRelsPath = f'{self._wordDir}/_rels/document.xml.rels'
        self._contentTypesPath = f'{self._tmpDir}/[Content_Types].xml'

    def open(self):
        with zipfile.ZipFile(self._docPath, 'r') as zipFile:
            zipFile.extractall(f'{self._tmpDir}')

    def save(self):
        shutil.make_archive(self._docPath, 'zip', self._tmpDir)
        p = pathlib.Path(f'{self._docPath}.zip')
        p.rename(pathlib.Path(p.parent, f'{self._docPath}'))

    def getHeaders(self):
        headers = []
        wordDocRoot = lxml.etree.parse(self._wordDocPath)
        pgMar = wordDocRoot.find('.//w:pgMar', DocXType.xmlns)
        print(f'[DocX.getHeaders] w:pgMar found: {pgMar}')
        print(pgMar.get(QName('w', 'left').text))
        print(pgMar.get(QName('w', 'header').text))
        srcRoot = lxml.etree.parse(self._wordRelsPath).getroot()
        for child in srcRoot:
            # print(f'[DocX.getHeaders] child: {lxml.etree.tostring(child, pretty_print=True)}')
            if 'Type' in child.keys():
                if child.attrib['Type'] == DocXType.header:
                    if 'Target' in child.keys():
                        headerFileName = child.attrib['Target']
                        print(f'[DocX.getHeaders] header docRels found: {headerFileName}')
                        headerXml: lxml.etree.ElementTree = lxml.etree.parse(f'{self._wordDir}/{headerFileName}')
                        headers.append(
                            {
                                'docRels': child,
                                'headerXmlName': headerFileName,
                                'headerXmlContent': headerXml,
                                'sectPrRef': None,
                                'w:sectPr.w:pgMar.left': pgMar.get(QName('w', 'left').text) if pgMar is not None else None,
                                'w:sectPr.w:pgMar.right': pgMar.get(QName('w', 'right').text) if pgMar is not None else None,
                                'w:sectPr.w:pgMar.gutter': pgMar.get(QName('w', 'gutter').text) if pgMar is not None else None,
                                'w:sectPr.w:pgMar.footer': pgMar.get(QName('w', 'footer').text) if pgMar is not None else None,
                                'w:sectPr.w:pgMar.bottom': pgMar.get(QName('w', 'bottom').text) if pgMar is not None else None,
                                'w:sectPr.w:pgMar.header': pgMar.get(QName('w', 'header').text) if pgMar is not None else None,
                                'w:sectPr.w:pgMar.top': pgMar.get(QName('w', 'top').text) if pgMar is not None else None,
                            }
                        )
        return headers

    def addHeader(self, srcHeader: dict):
        print(f'[DocX.addHeader] adding header: {srcHeader}')
        wordRelsRoot = lxml.etree.parse(self._wordRelsPath).getroot()
        headerId = self._docRelsAppendHeader(wordRelsRoot, srcHeader['docRels'])
        wordRelsRoot.getroottree().write(self._wordRelsPath, xml_declaration=True, encoding="utf-8")
        headerXml: lxml.etree.ElementTree = srcHeader['headerXmlContent']
        # wordRelsRoot = lxml.etree.open(f'{self._wordDir}/{srcHeader["headerXmlName"]}', mode='w').getroot()
        headerXml.write(f'{self._wordDir}/{srcHeader["headerXmlName"]}', xml_declaration=True, encoding="utf-8", standalone=True, pretty_print=True)
        wordDocRoot = lxml.etree.parse(self._wordDocPath).getroot()
        sectPr = wordDocRoot.find('.//w:sectPr', DocXType.xmlns)
        sectPr.insert(0, lxml.etree.fromstring(f'<w:headerReference xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="{headerId}" />'))
        pgMar = wordDocRoot.find('.//w:pgMar', DocXType.xmlns)
        # pgMar.attrib[QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'header').text] = "0"
        # pgMar.attrib[QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'top').text] = "1134"
        pgMar.set(QName('w', 'header').text, '0')
        pgMar.set(QName('w', 'top').text, '1134')
        wordDocRoot.getroottree().write(self._wordDocPath, xml_declaration=True, encoding="utf-8", standalone=True, pretty_print=True)
        # w:sectPr
        # <w:headerReference w:type="default" r:id="rId2" />
        self._updateContentTypes(srcHeader["headerXmlName"])

    def _updateContentTypes(self, fileName: str):
        # <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" />
        contentTypes = lxml.etree.parse(self._contentTypesPath).getroot()
        updated = False
        print(f'[DocX.addHeader] looking for: {fileName}')
        for child in contentTypes:
            # print(f'[DocX.addHeader] child.attrib.keys(): {child.attrib.keys()}')
            if 'PartName' in child.attrib.keys():
                # print(f'[DocX.addHeader] verifying: {child.attrib["PartName"]}')
                if child.attrib['PartName'] == f'/word/{fileName}':
                    # print(f'[DocX.addHeader] found: {child}')
                    updated = True
                    child.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
        if not updated:
            contentTypes.append(
                lxml.etree.fromstring(f'<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" />')
            )
        contentTypes.getroottree().write(self._contentTypesPath, xml_declaration=True, encoding="utf-8", standalone=True, pretty_print=True)

    def _getMaxId(self, root, attr: str):
        idMax = 1
        idPrefix = ''
        for child in root:
            id = child.attrib[attr]
            idParts = re.match(r'(\D+)(\d+)', id, re.I).groups()
            # print(f'[DocX._getMaxId] idParts: {idParts}')
            idPrefix = idParts[0]
            idSufix = int(idParts[1])
            # print(f'[DocX._getMaxId] idSufix: {idSufix}')
            if idSufix >= idMax:
                idMax = idSufix + 1
        # print(f'[DocX._getMaxId] idPrefix: {idPrefix}')
        # print(f'[DocX._getMaxId] idMax: {idMax}')
        return f'{idPrefix}{idMax}'

    def _docRelsAppendHeader(self, dstRoot, header):
        headerId = self._getMaxId(root = dstRoot, attr = 'Id')
        header.set('Id', headerId)
        dstRoot.append(header)
        return headerId



if __name__ == '__main__':
    path = pathlib.Path(__file__).parent.resolve()
    srcPath = f'{path}/source.docx'
    dstPath = f'{path}/target.docx'

    srcDoc = DocX(srcPath)
    dstDoc = DocX(dstPath)
    srcDoc.open()
    dstDoc.open()
    # exit(0)
    srcHeaders = srcDoc.getHeaders()
    for srcHeader in srcHeaders:
        print(f'[main] srcHeader: {srcHeader}')
        dstDoc.addHeader(srcHeader)

    dstDoc.save()
