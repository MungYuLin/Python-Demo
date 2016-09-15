# -*- coding: utf-8 -*-

from zipfile import ZipFile
from win32com.client import Dispatch
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML

ROOT_PATH = 'E:\\Python\\'
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'

def xml2text(text):
    tree = XML(text)
    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))
    return '\n\n'.join(paragraphs)

def docx2xml(file):
    zf = ZipFile(file)
    if not zf.namelist().__contains__('word/document.xml'):
        print '无效的MS文件.'
        return None
    text = str(zf.read('word/document.xml'))
    zf.close()
    return xml2text(text)

def doc2html(file):
    '''
        通过python启动办公软件的应用进程
        其中wps、et、wpp对应的是金山文件、表格和演示
        word、excel、powerpoint对应的是微软的文字、表格和演示
        天知道为啥用金山的打不开，用为微软的打开了金山的wps
    '''
    lst_app_name = [
        'wps.Application',
        'et.Application',
        'wpp.Application',
        'word.Application',
        'excel.Application',
        'powerpoint.Application'
    ]
    wordFullName = ROOT_PATH + file
    dotIndex = wordFullName.find(".")
    if (dotIndex == -1):
        print '未取得后缀名.'
        return None
    fileSuffix = wordFullName[(dotIndex + 1):]
    if (fileSuffix == "doc" or fileSuffix == "docx"):
        htmlFullName = wordFullName[:dotIndex] + '.html'
        wpsApp = Dispatch('word.application')
        doc = wpsApp.Documents.Open(wordFullName)
        doc.SaveAs(htmlFullName, 10)
        doc.Close()
        wpsApp.Quit()
    else
        print '不是文档'
    return

if __name__ == '__main__':
    doc2html('test.doc')
    # print(docx2xml('test.docx'))
