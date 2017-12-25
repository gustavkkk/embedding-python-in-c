# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 17:13:08 2017

@author: Frank
"""

from processing.context import fill_in_paragraph,preprocess,find_keywords
from processing.table import fill_in_table,get_tbl_name
from processing.picture import insert_pictures
from processing.tag import check_tag_is,get_text
#from lxml import etree
import os
from openxml import OpenXML
#from PyQt5.QtWidgets import QProgressBar

class FillIn(OpenXML):
    
    def __init__(self,input_filepath,tmpdir=None):
        super(FillIn, self).__init__(input_filepath)
        self.section_name = os.path.splitext(os.path.basename(input_filepath))[0]
        if tmpdir is not None:
            self.tmpdir = tmpdir
        else:
            self.tmpdir = os.path.join(os.getcwd(),'tmp/split-processed')
        
    def process(self,output_filepath,progressbar=None):
        self.preprocess(progressbar)
        self.process_paras(progressbar)
        self.process_tbls(progressbar)
        self.save(output_filepath)
        self.insert_imgs(filepath=output_filepath)

    def preprocess(self,progressbar=None):
        # init progressbar
        if progressbar:
            progressbar.setValue(0)
        #
        p_count = len(self.body)
        for i,p in enumerate(self.body):
            if progressbar:
                progressbar.setValue(int((i+1)*100/p_count))
            if check_tag_is(p,'w:p'):
                text = get_text(p)
                if text == "":
                    continue
                print(text)
                preprocess(p)        
    # filling-in should be automatically, smartly progressed.
    # section-name analysis, word analysis, sentence analysis, co-related words analysis(table analysis) are in need
    # fixed input should be avoided. everything must be flexible.
    # the present implementation is just primitive, pretty fixed.
    # word similarity analysis,etc should be help
    # <w:p w:rsidR="00486677" w:rsidRDefault="00486677">
    #   <w:pPr>
    #       <w:tabs/>
    #       <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    #-------<w:ind w:left="274" w:right="-20"/>--------------------margin
    #       <w:rPr/>
    #   </w:pPr>
    #   <w:r/>
    #   <w:r>
    #       <w:rPr>
    #           <w:rFonts/>
    #           <w:sz />
    #           <w:szCs/>
    #-----------<w:u/>-------underline
    #           <w:lang/>
    #-------<w:t/>-----------text
    #-------<w:tab/>---------tab
    # </w:p>
    def process_paras(self,progressbar=None):
        # init progressbar
        if progressbar:
            progressbar.setValue(0)
        #
        p_count = len(self.body)
        for i,p in enumerate(self.body):
            if progressbar:
                progressbar.setValue(int((i+1)*100/p_count))
            if check_tag_is(p,'w:p'):
                text = get_text(p)
                if text == "":
                    continue
                itemlist = find_keywords(text)
                fill_in_paragraph(self.section_name,p,itemlist)
    # filling-in should be automatically, smartly progressed.
    # section-name analysis, word analysis, sentence analysis, corelated words analysis(table analysis) are in need
    # fixed input should be avoided. everything must be flexible.    
    # the present implementation is just primitive, pretty fixed.
    # word similarity analysis,etc should be help
    # <w:tbl>
    #   <w:tblPr>...</w:tblPr>
    #       <w:tblW w:w="0" w:type="auto"/>
    #       <w:jc w:val="center"/>
    #       <w:tblLayout w:type="fixed"/>
    #       <w:tblCellMar>
    #           <w:left w:w="0" w:type="dxa"/>
    #           <w:right w:w="0" w:type="dxa"/>
    #       </w:tblCellMar>
    #       <w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>
    #   <w:tblGrid>
    #-------<w:gridCol w:w="891"/>|
    #-------<w:gridCol w:w="891"/>|—————————————————————————col count
    #-------<w:gridCol w:w="891"/>|
    #       ...
    #   </w:tblGrid>
    #   <w:tr w:rsidR="00486677">...</w:tr>
    #       <w:trPr>...</w:trPr>
    #       <w:tc>...</w:tc>
    #           <w:tcPr>
    #               <w:tcW w:w="5950" w:type="dxa"/>
    #---------------<w:vMerge w:val="restart"/>—————————————rows merged
    #---------------<w:gridSpan w:val="5"/>—————————————————cols merged
    #               <w:tcBorders>...</w:tcBorders>
    #               <w:shd w:val="clear" w:color="000000" w:fill="FFFFFF"/>
    #               <w:vAlign w:val="center"/>
    #           <w:p w:rsidR="00486677" w:rsidRDefault="00486677">...</w:p>
    # </w:tbl>
    def process_tbls(self,progressbar=None):
        # init progressbar
        if progressbar:
            progressbar.setValue(0)
        #
        p_count = len(self.body)
        for i,tbl in enumerate(self.body):
            if progressbar:
                progressbar.setValue(int((i+1)*100/p_count))
            if check_tag_is(tbl,'w:tbl'):
                print('processing tbl')
                fill_in_table(self.section_name,tbl)
    
    def insert_imgs(self,filepath):
        insert_pictures(self.section_name,filepath)

tmp_dir1 = os.path.join(os.getcwd(),'tmp\\split\\')
tmp_dir2 = os.path.join(os.getcwd(),'tmp\\split-processed\\')

from config import INDEX
import glob

def main():
    filepaths = glob.glob(tmp_dir1+'*.docx')
    sections = [os.path.basename(filepath)[:os.path.basename(filepath).find(".docx")] for filepath in filepaths]
    dic = {}
    for section in sections:
        for index in list(INDEX.keys()):
            if index in section:
                dic[section] = INDEX.index(index)
                break
    sections = sorted(dic, key=dic.__getitem__)
    print(sections)
    #
    for section in sections:
        fillin = FillIn(os.path.join(tmp_dir1,section+'.docx'))
        fillin.process(os.path.join(tmp_dir2,section+'.docx'))

from processing.convert import docx2xml

def test(target='授权委托书'):
    from config import db,POSITION
    for position in POSITION:
        if '项目经理' in position:
            db['project_members'][position] = '秦仁和'
        else:
            db['project_members'][position] = '王养'
    
    db['human'].select_people(name_list=db['project_members'].values())
    for filepath in glob.glob(tmp_dir1+'*.docx'):
        if target in filepath:
            break
    print(filepath)
    fillin = FillIn(filepath)
    docx_outpath = os.path.join(os.getcwd(),target+'.docx')
    fillin.process(docx_outpath)
    docx2xml(docx_outpath,os.path.join(os.getcwd(),target+'.xml'))
    
if __name__ == "__main__":
    test()