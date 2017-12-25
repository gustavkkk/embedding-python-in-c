# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 15:12:09 2017

@author: Frank
"""

from zipfile import ZipFile
from lxml import etree
import os
import shutil
import copy

filenames = ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/document.xml', 'word/theme/theme1.xml', 'word/settings.xml', 'word/fontTable.xml', 'word/webSettings.xml', 'docProps/app.xml', 'docProps/core.xml', 'word/styles.xml']
    
def clone(tree):
    return copy.deepcopy(tree)#etree.fromstring(etree.tostring(tree))#

def docx2xml(docx_filename,xml_filename=None):
    with open(docx_filename,'r+b') as docx_file:
        xml = ZipFile(docx_file).read('word/document.xml')
        xmltree = etree.fromstring(xml)
        prettyxml = etree.tostring(xmltree,pretty_print=True)
        if xml_filename is None:
            xml_filename =  os.path.splitext(docx_filename)[0] + '.xml'
        with open(xml_filename,'w+b') as xml_file:
            xml_file.write(prettyxml)

def docx2xmltree(docx_filename):
    with open(docx_filename,'r+b') as docx_file:
        xml = ZipFile(docx_file).read('word/document.xml')
        xmltree = etree.fromstring(xml)
        return xmltree

def save_xml_tree(xml_tree,xml_filename):
    with open(xml_filename,'w+b') as xml_file:
        xml_string = etree.tostring(xml_tree,pretty_print=True)
        xml_file.write(xml_string)
            
def unzip(zipfilename, dest_dir):
    if not os.path.exists(os.path.abspath(zipfilename)):
        return
    if not os.path.exists(dest_dir):
        os.mkdir(dest_dir)
    with ZipFile(zipfilename) as zf:
        zf.extractall(dest_dir)

def xmltree2docx(xmltree,docx_filename):
    prettyxml = etree.tostring(xmltree,pretty_print=True)
    with open(os.path.join(os.getcwd(),'tmp/word/document.xml'),'w+b') as xml_file:
        xml_file.write(prettyxml)
    #
    with ZipFile(docx_filename, "w") as docx_:
        for filename in filenames:
            docx_.write(os.path.join(os.getcwd(),'tmp/'+filename), filename)       
    
def xml2docx(xml_filename,docx_filename):
    shutil.copy(xml_filename, os.path.join(os.getcwd(),'tmp/word/document.xml'))
    #
    with ZipFile(docx_filename, "w") as docx_:
        for filename in filenames:
            docx_.write(os.path.join(os.getcwd(),'tmp/'+filename), filename)


import subprocess
import comtypes.client
from win32com.client import Dispatch

WORD = 'Word.Application'
wdFormatPDF = 17#https://msdn.microsoft.com/en-us/VBA/Word-VBA/articles/wdsaveformat-enumeration-word
wdFormatDocx = 12

def opendocument(filepath):
    word = Dispatch(WORD)
    word.Visible = True
    word.DisplayAlerts = 0    
    word.Documents.Open(filepath)
    #doc.Close()
    #word.Quit()
    
def doc2docx_under_test(doc_filepath,docx_filepath=None):
    subprocess.call('soffice --headless --convert-to docx'.format(doc_filepath), shell=True)#(['soffice', '--headless', '--convert-to', 'docx', doc_filepath])

def doc2docx(doc_filepath,docx_filepath=None):
    #https://social.msdn.microsoft.com/Forums/en-US/a4f00910-cb6e-4861-bf96-97b0cfc6cf8f/convert-word-files-from-doc-to-docx-using-python?forum=worddev
    word = Dispatch(WORD)
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(doc_filepath)
    if docx_filepath is None:
        docx_filepath = doc_filepath.replace(".doc", r".docx")
    #print(docx_filepath)
    doc.SaveAs2(docx_filepath, FileFormat=wdFormatDocx)
    doc.Close()
    word.Quit()
    
def docx2pdf(docx_filepath,pdf_filepath):
    word = comtypes.client.CreateObject(WORD)
    word.Visible = False
    doc = word.Documents.Open(docx_filepath)
    doc.SaveAs2(pdf_filepath, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

from pydocx import PyDocX
def docx2html(docx_filepath,html_filename=None):
    with open(docx_filepath, 'rb') as docx_file:
        html = PyDocX.to_html(docx_file)
        xmltree = etree.fromstring(html)
        prettyxml = etree.tostring(xmltree,pretty_print=True)
        if html_filename is None:
            html_filename =  os.path.splitext(docx_filepath)[0] + '.html'
        with open(html_filename,'w+b') as html_file:
            html_file.write(prettyxml)

import docx
def docx2text(docx_filepath):
    doc = docx.Document(docx_filepath)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

def pdf2docx_under_test(pdf_filepath,docx_filepath=''):
    subprocess.call('lowriter --invisible --convert-to doc "{}"'
                    .format(pdf_filepath), shell=True)

def pdf2docx_dir_under_test(path='/my/pdf/folder'):
    for top, dirs, files in os.walk(path):
        for filename in files:
            if filename.endswith('.pdf'):
                abspath = os.path.join(top, filename)
                subprocess.call('lowriter --invisible --convert-to doc "{}"'
                                .format(abspath), shell=True)

def pdf2docx(pdf_filepath,docx_filepath=None):
    if not os.path.exists(pdf_filepath):
        return
    word = Dispatch(WORD)
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(pdf_filepath)
    if docx_filepath is None:
        docx_filepath = pdf_filepath.replace(".pdf", r".docx")
    doc.SaveAs2(docx_filepath, FileFormat=wdFormatDocx)
    doc.Close()
    word.Quit()    
    
#pip install pypdf2
#https://www.lfd.uci.edu/~gohlke/pythonlibs/#pythonmagick
#https://stackoverflow.com/questions/13984357/pythonmagick-cant-find-my-pdf-files
#https://www.ghostscript.com/download/gsdnld.html
#from PyPDF2 import PdfFileReader
from PythonMagick import Image
def pdf2image(pdf_filepath):
    if not os.path.exists(pdf_filepath):
        return    
    '''
    pdf = PdfFileReader(pdf_filepath)
    for page_number in range(pdf.getNumPages()):
        image = Image(pdf.getPage(page_number+1))
        image.write('file_image{}.png'.format(page_number+1))
    pass
    '''
    im = Image(pdf_filepath)
    im.write("file_img%d.png")

pdf_filepath = os.path.join(os.getcwd(),'..\\..\\Documentation\\test.pdf')

def test():
    pdf2docx(pdf_filepath)
    #print(pdf_filepath)
    #pdf2docx_under_test(pdf_filepath)
    '''
    if os.path.exists(pdf_filepath):
        print(pdf_filepath)
        pdf2image(pdf_filepath)
    '''    
    
if __name__ == "__main__":
    unzip('..\\pic.docx','..\\test')
    