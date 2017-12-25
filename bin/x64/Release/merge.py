# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 18:02:05 2017

@author: Frank
"""
import os
#import shutil
import copy        
import glob
from docx import Document
#from lxml import etree
from processing.convert import docx2xmltree,xmltree2docx,save_xml_tree
#from tagprocessing import del_attrib_pgNumType
class Merge:
    
    def __init__(self,tmpdir=None,section_names=None):
        if tmpdir is None:
            self.tmpdir = os.path.join(os.getcwd(),'tmp/split/')
        else:
            self.tmpdir = tmpdir
        if section_names is None:
            self.files = glob.glob(self.tmpdir+'*.docx')
        else:
            self.files = [os.path.join(self.tmpdir,section_name+'.docx') for section_name in section_names]
    #
    def process(self,fullpath):
        self.process_by_docx(fullpath)
        
    def process_by_docx(self,fullpath):
        combined_document = Document()
        count = 0
        for file in self.files:
            sub_doc = Document(file)    
            # Don't add a page break if you've
            # reached the last file.
            #if count < len(files) - 1:
            #    sub_doc.add_page_break()
    
            for element in sub_doc.element.body:
                combined_document.element.body.append(element)
            count += 1
        combined_document.save(fullpath)    
    #
    def process_by_lxml(self,fullpath):
        xtree = docx2xmltree(self.files[0])
        xbody = xtree[0]
        for i in range(1,len(self.files)):
            xtree_ = docx2xmltree(self.files[i])
            xbody_ = xtree_[0]
            for child in xbody_:
                xbody[len(xbody)-1].addnext(copy.deepcopy(child))
        xmltree2docx(xtree,fullpath)
        #save_xml_tree(xtree,os.path.join(os.path.split(fullpath)[0],'merge.xml'))
    #
        
output_filename = 'output.docx'
fullpath = os.path.join(os.getcwd(),output_filename)

from processing.convert import xml2docx

def testxml():
    xml2docx(os.path.join(os.getcwd(),'merge.xml'),fullpath)
    
if __name__ == "__main__":
    merge = Merge()
    merge.process(fullpath)
    #testxml()