# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 15:58:22 2017

@author: Frank
"""
import os
import copy        
from processing.tag import check_tag_is,get_text
from processing.context import segment
from processing.convert import docx2xmltree,xmltree2docx
from processing.util import extract_pages,split_page_by_brackets
from config import INDEX
import numpy as np
from misc import resource_path,mkdir
#from PyQt5.QtWidgets import QProgressBar

import jieba
#http://blog.csdn.net/qq_26376175/article/details/69680992
jieba.set_dictionary("./dict.txt")  
jieba.initialize()
#from jieba import analyse  
#jieba.analyse.set_idf_path("./idf.txt")  

for word in list(INDEX.keys()):
    jieba.add_word(word)

class Split:
    
    def __init__(self,input_filepath,tmpdir=None):
        if tmpdir is None:
            self.tmpdir = os.path.join(os.getcwd(),'tmp\\split\\')
        else:
            self.tmpdir = tmpdir
        mkdir(resource_path(self.tmpdir))
        self.xml_tree = docx2xmltree(input_filepath)
        self.xml_body = self.xml_tree[0]
    
    def make_index_list(self,keyword="目录"):
        #
        dic = {}
        # make indexlist
        indexlist = []
        for i,child in enumerate(self.xml_body):
            if check_tag_is(child,'w:p'):
                text = get_text(child)
                if text == "":
                    continue
                seglist = segment(text)
                if keyword not in seglist:
                    continue
                
                dic[keyword] = i
                for count in range(30):
                    child = child.getnext()
                    text = get_text(child)
                    if text == "":
                        continue
                    seglist = segment(text)
                    #print(seglist)
                    seglist.sort(key=lambda string:len(string),reverse=True)
                    #print(seglist[0])
                    #print(text)
                    if seglist[0] not in list(INDEX.keys()) or (len(indexlist) > 0 and seglist[0] in list(np.array(indexlist)[:,1])):
                        break
                    else:
                        indexlist.append([text,seglist[0]])
                break
        #print(indexlist)
        #print(dic)
        #print('split:make_index_list:indexlist\n',indexlist)
        #print(indexlist)
        # save page of subsystem
        liter = iter(indexlist)
        [heading,key] = next(liter)
        keys = [key]
        dic_init_size = len(dic)
        for j,p in enumerate(self.xml_body):
            if j < i+count+1:
                continue
            if check_tag_is(p,'w:p'):
                text = get_text(p)
                if text == "":
                    continue
                if any(key in text for key in keys) and len(text) < len(heading)*2:#any(index in text for index in indexlist):
                    #print(j,text,keys,heading)
                    dic[heading] = j
                    if (len(dic) - dic_init_size) == len(indexlist):
                        break
                    [heading,key] = next(liter)
                    keys = []
                    keys.append(key)
                    if key == "其他资料":
                        keys.append("其他材料")
        #print('split:dic:\n',dic)
        #print('split:indexlist:\n',indexlist)
        heading_titles = sorted(dic, key=dic.__getitem__)
        #heading_pages = sorted(dic.values())
        return [[title,dic[title]] for title in heading_titles]
    
    def process(self,progressbar=None):
        # init progressbar
        if progressbar:
            progressbar.setValue(0)
        #
        sections = self.make_index_list()
        #print(sections)
        sections.insert(0,['封面页',0])
        section_count = len(sections)
        sections_all = []
        #print('split:process:sections:\n',sections)
        for index,section in enumerate(sections):
            section_name,first_para_index = section
            xtree = copy.deepcopy(self.xml_tree)
            xbody = xtree[0]
            if index == len(sections) - 1:
                last_para_index = len(xbody) - 1
            else:
                last_para_index = sections[index+1][1] - 1
            extract_pages(xbody,first_para_index,last_para_index)
            if any(key in section_name for key in ['目录','已标价工程量清单']):
                xmltree2docx(xtree,resource_path(os.path.join(os.getcwd(),self.tmpdir+section_name+'.docx')))
                sections_all.append(section_name)
            else:
                for xtree_,section_name_ in split_page_by_brackets(xtree,section_name):
                    filepath = resource_path(os.path.join(os.getcwd(),self.tmpdir+section_name_+'.docx'))
                    print(filepath)
                    xmltree2docx(xtree_,filepath)
                    sections_all.append(section_name_)
            if progressbar:
                progressbar.setValue(int((index+1)*100/section_count))
        #            
        return sections_all#list(np.array(sections)[:,0])

    '''    
    def process_(self,progressbar=None):
        # init progressbar
        if progressbar:
            progressbar.setValue(0)
        #
        sections = self.make_index_list()
        sections.insert(0,['封面页',0])
        section_count = len(sections)
        #print('split:process:sections:\n',sections)
        for index,section in enumerate(sections):
            section_name,first_para_index = section
            xtree = copy.deepcopy(self.xml_tree)
            xbody = xtree[0]
            if index == len(sections) - 1:
                last_para_index = len(xbody) - 1
            else:
                last_para_index = sections[index+1][1] - 1
            extract_pages(xbody,first_para_index,last_para_index)
            xmltree2docx(xtree,os.path.join(os.getcwd(),self.tmpdir+section_name+'.docx'))
            if progressbar:
                progressbar.setValue(int((index+1)*100/section_count))
        return list(np.array(sections)[:,0])
    '''

input_filename = 'simohua-extract.docx'
fullpath = resource_path(os.path.join(os.getcwd(),input_filename))

import sys

def split(filepath=fullpath):
    split = Split(filepath)
    sections = split.process()
    print(sections)
    return sections

'''
python extract.py simohua.docx
'''    
if __name__ == "__main__":
    filepath = resource_path(sys.argv[1])
    if not os.path.exists(filepath):
        print('no such a file exists')
    split(filepath)
