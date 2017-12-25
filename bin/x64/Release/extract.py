# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 16:58:45 2017

@author: Frank
"""

import os
#import shutil
#import copy
#from lxml import etree

from config import KEYWORDS,db,SYNONYMS,PROJECT_INFO_KEY
from processing.tag import check_tag_is,get_text,del_tags_by_name,remove_all_childs_before_index
from processing.context import extract_info_by_colon
from processing.table import extract_project_info_from_tbl
from misc import resource_path
from openxml import OpenXML

import jieba
#http://blog.csdn.net/qq_26376175/article/details/69680992
jieba.set_dictionary("./dict.txt")  
jieba.initialize()
#from jieba import analyse  
#jieba.analyse.set_idf_path("./idf.txt")  

for word in KEYWORDS:
    jieba.add_word(word)

def segment(text):
    seglist = jieba.cut(text, cut_all=False)
    return list(seglist)

def get_synonyms(word):
    for cluster in SYNONYMS:
        if word in cluster:
            return cluster
    return [word]

def set_output_filepath(input_filepath,attach=''):
    #basename = os.path.basename(input_filepath)
    #output_filename = os.path.splitext(basename)[0] +\
    #                  'extracted' + \
    #                  os.path.splitext(basename)[1]
    return os.path.splitext(input_filepath)[0] + attach + os.path.splitext(input_filepath)[1]
    
class Extract(OpenXML):
    
    def __init__(self,input_filepath):
        super(Extract, self).__init__(input_filepath)
        self.output_filepath = os.path.splitext(input_filepath)[0] + '-extract' + os.path.splitext(input_filepath)[1]
        
    def del_page_num_tag(self):
        for child in self.body:
            del_tags_by_name(child)            
            
    def extract_cover_page_info(self,keyword="投标文件格式"):
        # find the position of coverpage
        # in’投标文件格式‘
        # out:‘第九章’
        #index = -1
        for child in self.body:
            # only paragraph
            if check_tag_is(child,'w:p'):
                text = get_text(child)
                if text == "":
                    continue
                seglist = segment(text)
                # only with keyword
                for key_word in get_synonyms(keyword):
                    if key_word in seglist:
                        #print(seglist)
                        idx = seglist.index(key_word)
                        #position = seglist[index+2] if index+2 < len(seglist) else seglist[index+1]
                        pos_idx = idx# + 2
                        while pos_idx < len(seglist) and (57 < ord(seglist[pos_idx][0]) or ord(seglist[pos_idx][0]) < 48):
                            pos_idx += 1
                        return seglist[idx-1],key_word,seglist[pos_idx]
        return None,None,None
 
    # This function is constant now
    # It should be replaced with real-extractions from docx
    def extract_project_infos(self):
        indexes = [self.find_cover_page_para_index(keyword=keyword) for keyword in ['招标公告','投标人须知','评标办法']]
        match = {}
        print(indexes)
        for index in indexes:
            if index == -1:
                print('找不到招标公告及招标人须知')
                match['招标人名称'] = '巴东县水土保持局'
                match['招标人'] = '巴东县水土保持局'
                match['项目名称'] = '巴东县2017年岩溶地区石漠化综合治理工程'
                return match
        project_name_set = False
        for i,child in enumerate(self.body):
            #if i < indexes[0]:
            if not project_name_set:
                if not check_tag_is(child,'w:p'):# or project_name_set:
                    continue
                text = get_text(child)
                if text == '':
                    continue
                match['项目名称'] = text
                project_name_set = True
            #elif indexes[0] <= i and i < indexes[1]:
            elif 2 < i and i < indexes[1]:
                if not check_tag_is(child,'w:p'):
                    continue
                text = get_text(child)
                if text == '':
                    continue
                ind = extract_info_by_colon(text)
                print(ind)
                for key in list(ind.keys()):
                    if key in PROJECT_INFO_KEY:
                        if key not in list(match.keys()):
                            match[key] = ind[key]
                        elif len(match[key]) > len(ind[key]):
                            match[key] = ind[key]
                            
            elif indexes[1] <= i and i < indexes[2]:
                if not check_tag_is(child,'w:tbl'):
                    continue
                ind = extract_project_info_from_tbl(child)
                for key in list(ind.keys()):
                    if key in PROJECT_INFO_KEY:
                        if key not in list(match.keys()):
                            match[key] = ind[key]
                        elif len(match[key]) > len(ind[key]):
                            match[key] = ind[key]
                            
                #if len(ind) > 1:
                #    print(match)
                #    return match
            elif indexes[2] <= i:
                break
        #print(match)
        return match
    
    def find_cover_page_para_index(self,keyword='投标文件格式'):
        chapter_number,key_word,position = self.extract_cover_page_info(keyword=keyword)
        if chapter_number is None:
            return -1
        # toss
        avg_paras_per_page = 10
        estimated_position = int(position) * avg_paras_per_page
        
        # find paragraph index of coverpage
        # in: ‘第九章’，‘投标文件格式’
        # out: 2320(self.body[2320] is what)
        for i,child in enumerate(self.body):
             if i < estimated_position or not check_tag_is(child,'w:p'):
                 continue
             text = get_text(child)
             if text == "":
                 continue
             seglist = segment(text)
             if all(key in seglist for key in [chapter_number,key_word]):
                 if abs(seglist.index(chapter_number) - seglist.index(key_word)) < 2:
                     #print(i,seglist)
                     return i             
        return -1

    def process(self,output_filepath=None):
        #
        db['project_info'].set_db(self.extract_project_infos())
        #
        self.remove_all_pages_before_cover_page()
        self.del_page_num_tag()
        if output_filepath is None:
            output_filepath = self.output_filepath
        self.save(resource_path(output_filepath))
        
    def remove_all_pages_before_cover_page(self):
        #
        coverpage_index = self.find_cover_page_para_index()
        #print(coverpage_index)
        remove_all_childs_before_index(self.body,coverpage_index)

            
fullcontent_filename = 'simohua.docx'
fullpath = resource_path(os.path.join(os.getcwd(),fullcontent_filename))

import sys

def extract(filepath=fullpath):
    extract = Extract(filepath)
    extract.process()

'''
python extract.py simohua.docx
'''    
if __name__ == "__main__":
    filepath = resource_path(sys.argv[1])
    if not os.path.exists(filepath):
        print('no such a file exists')
    extract(filepath)
