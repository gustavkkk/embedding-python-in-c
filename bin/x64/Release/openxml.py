# -*- coding: utf-8 -*-
"""
Created on Mon Nov 20 00:26:01 2017

@author: Frank
"""
import os
from zipfile import ZipFile
from lxml import etree
import tempfile
import shutil
from docx import Document
import datetime

class OpenXML:
    
    def __init__(self,input_filepath):
        self.file = open(input_filepath,'r+b')
        self.zipfile = ZipFile(self.file)
        self.xml_content = self.zipfile.read('word/document.xml')
        self.xml_tree = etree.fromstring(self.xml_content)
        self.body = self.xml_tree[0]
        self.input_filepath = input_filepath
        self.date = datetime.datetime.now()
        
    def process(self):
        self.zipfile.close()
        self.file.close()
        pass
    
    def save(self,output_filepath):
        #in:xml_tree
        #out:.docx
        tmp_dir = tempfile.mkdtemp()
        self.zipfile.extractall(tmp_dir)
        with open(os.path.join(tmp_dir,'word/document.xml'), 'w+b') as f:
            xmlstr = etree.tostring(self.xml_tree, pretty_print=True)
            f.write(xmlstr)
        # Get a list of all the files in the original docx zipfile
        filenames = self.zipfile.namelist()
        # Now, create the new zip file and add all the filex into the archive
        zip_copy_filepath = output_filepath
        with ZipFile(zip_copy_filepath, "w") as docx_:
            for filename in filenames:
                docx_.write(os.path.join(tmp_dir,filename), filename)    
        # Clean up the temp dir
        shutil.rmtree(tmp_dir)
        #
        self.zipfile.close()
        self.file.close()
        #
        doc = Document(output_filepath)
        doc.save(output_filepath)