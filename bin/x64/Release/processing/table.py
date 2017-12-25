# -*- coding: utf-8 -*-
"""
Created on Mon Dec  4 15:58:27 2017

@author: Frank
"""
from .tag import check_tag_is,check_tag_exist
from .tag import get_text
from .tag import get_attrib_value_word
from misc import switch

#classification
#      _______________________________________________
#      |       |            |           |             |
#      |  no   |   name     |  content  |    remake   |——all-black
#      |_______|____________|___________|_____________|
#      |_______|____________|___________|_____________|——all-white
#      |_______|____________|___________|_____________|——all-white
#      |       |            |           |             |
#      | name  |            |  position |             |——zebra
#      |_______|____________|___________|_____________|
#      |       |            |           |             |
#      | sex   |            | telephone |             |——zebra
#      |_______|____________|___________|_____________|
#      |       |        |          |        |         |
#      | title |  unit  |   2014   |  2015  |  2016   |——all-black
#      |_______|________|__________|________|_________|
#      |       |        |          |        |         |
#      | profit|        |          |        |         |——one-black
#      |_______|________|__________|________|_________|
#
def analyze_tr(tr):
    tc_count =  0#len(tr) - 1
    tc_count_text = 0
    for tc in tr:
        if check_tag_is(tc,'w:tc') or check_tag_exist(tc[0],'w:vMerge'):
            tc_count += 1
            if get_text(tc) is not '':
                tc_count_text += 1
    ''''''
#    if tc_count_text is tc_count:
#        return 'all black'#,cols,cols_
    if tc_count_text is 0:
        return 'all-white'#
    elif tc_count_text is 1 and tc_count > 2:
        return 'one-black' or 'all-white'#
    elif tc_count_text == tc_count:
        return 'all-black'#
    elif tc_count_text > tc_count * 0.5:
        return 'all-black'
    else:
        return 'zebra'
    ''''''

def analyze_tc(tc):
    if not check_tag_is(tc,'w:tc'):
        return -1
    for child in tc:
        if check_tag_is(child,'w:tcPr'):
            grid = check_tag_exist(child,'w:gridSpan')
            if grid is None:
                return 1
            else:
                cols = get_attrib_value_word(grid,'w:val')
                return int(cols)

import copy
from .tag import clear_text

def create_tc(tc,text):
    if not check_tag_is(tc,'w:tc'):
        return None
    tc_ = copy.deepcopy(tc)
    clear_text(tc_)
    isreplaced = False
    for p in tc_:
        if not check_tag_is(p,'w:p'):
            continue
        if isreplaced:
            break
        for r in p:
            if not check_tag_is(r,'w:r'):
                continue
            t = check_tag_exist(r,'w:t')
            if t is None:
                continue
            else:
                t.text = text
                isreplaced = True
                break
    return tc_   

def extract_project_info_from_tbls(body,tbl_names=['投标人须知前附表']):
    project_info_dic = {}
    for i,tbl in enumerate(body):
        if not check_tag_is(tbl,'w:tbl'):
            continue
        tbl_name = get_text(get_tbl_name(tbl))
        headers = get_tbl_header(tbl)
        if not tbl_name in tbl_names or not all(key in headers for key in ['条款号','条款名称','编列内容']):
            continue
        for tr in tbl:
            if not check_tag_is(tr,'w:tr'):
                continue
            col_index = 0
            keyword = ''
            for tc in tr:
                if not check_tag_is(tc,'w:tc'):
                    continue
                col_index += 1
                if col_index == 1:
                    continue
                fulltext = ''
                for p in tc:
                    if not check_tag_is(p,'w:p'):
                        continue
                    text = get_text(p)
                    if text == '':
                        continue
                    fulltext += text
                if col_index == 2:
                    keyword = fulltext
                elif col_index == 3:
                    project_info_dic[keyword] = fulltext
    return project_info_dic

from .context import extract_info_by_colon

def extract_project_info_from_tbl(tbl,tbl_names=['投标人须知前附表']):
    project_info_dic = {}
    #print('.')
    if not check_tag_is(tbl,'w:tbl'):
        return project_info_dic
    #print('..')
    tbl_name = get_text(get_tbl_name(tbl))
    #print(tbl_name)
    headers = get_tbl_header(tbl)
    #print(headers)
    if not tbl_name in tbl_names or not all(key in headers for key in ['条款号','条款名称','编列内容']):
        return project_info_dic
    #print('...')
    for tr in tbl:
        if not check_tag_is(tr,'w:tr'):
            continue
        col_index = 0
        keyword = ''
        for tc in tr:
            if not check_tag_is(tc,'w:tc'):
                continue
            col_index += 1
            if col_index == 1:
                continue
            fulltext = ''
            for p in tc:
                if not check_tag_is(p,'w:p'):
                    continue
                text = get_text(p)
                if text == '':
                    continue
                if any(key in text for key in ['：',':']):
                    ind = extract_info_by_colon(text)
                    # add keys
                    for key in list(ind.keys()):
                        if key not in list(project_info_dic.keys()):
                            if key in keyword:
                                project_info_dic[key] = ind[key]
                            elif col_index == 3:
                                project_info_dic[keyword+key] = ind[key]
                fulltext += text
            #print(keyword,fulltext)
            if col_index == 2:
                keyword = fulltext
            elif col_index == 3:
                project_info_dic[keyword] = fulltext
    return project_info_dic
        
from .tag import create_p

def fill_in_tc(tc,text):
    #print('...')
    if not check_tag_is(tc,'w:tc') or get_text(tc) is not '':
        return
    #print('....')
    p = check_tag_exist(tc,'w:p')
    p_x = create_p(fontsize=10)
    p_ = create_p(text)
    #print('.....')
    if p is None:
        print('fill_in_tc:creating paragraph')
        tc.insert(len(tc),p_x)
        p_x.addnext(p_)
        return
    else:
        print('fill_in_tc:replacing paragraph')
        p.addnext(p_x)
        p_x.addnext(p_)
        tc.remove(p)
    '''
    for p in tc:
        if not check_tag_is(p,'w:p'):
            continue
        for r in p:
            if not check_tag_is(r,'w:r'):
                continue
    '''
            
from .context import search_word_in_db,find_keywords,fill_in_paragraph,split_all_clusters
from config import POSITION

def fill_in_table(section_name,tbl):
    tbl_name_p = get_tbl_name(tbl)
    if tbl is None:
        tbl_name =  '无名'
    else:
        tbl_name =  get_text(tbl_name_p)
    parent = tbl.getparent()
    if tbl_name is None:
        print('表名称找不到了')
        return
    if any(key in section_name for key in ['管理机构','简历']):
        print(section_name)
        #position_list = list(db['project_members'].keys())[::-1]
        position_list = [position for position in POSITION if not position in ['法人','联系人','委托人1','委托人2']]
        if '管理机构' in tbl_name:
            #for keyword in keyword_list:
            #    name = db['project_members'][keyword]
            print('processing 管理机构')
            process_tbl(tbl,tbl_name,position_list)
        elif '简历' in tbl_name:
            last =  tbl
            print('processing 简历')
            for i,position in enumerate(position_list):
                #name = db['project_members'][keyword]
                #process_tbl(tbl,tbl_name,keyword)
                #set_tbl_name(tbl,tbl_name+'-'+str(i)+'('+keyword+')')
                para = copy.deepcopy(tbl_name_p)
                set_title(para,tbl_name+'-'+str(i)+'('+position+')')
                last.addnext(para)
                table = copy.deepcopy(tbl)
                process_tbl(table,tbl_name,position)
                para.addnext(table)
                #last = table
                #
                para2 = copy.deepcopy(tbl_name_p)
                set_title(para2,'附：'+position+'证件')#'附：参考图片')
                table.addnext(para2)
                last = para2
            parent.remove(tbl_name_p)
            parent.remove(tbl)
    elif '授权委托书' in section_name:
        for tr in tbl:
            if not check_tag_is(tr,'w:tr'):
                continue
            for tc in tr:
                if not check_tag_is(tc,'w:tc'):
                    continue
                for p in tc:
                    if not check_tag_is(p,'w:p'):
                        continue
                    text = get_text(p)
                    if text is '':
                        continue
                    keyword_list = find_keywords(text)
                    print(keyword_list)
                    split_all_clusters(p)
                    fill_in_paragraph(section_name,p,keyword_list)
                
    elif '投标函' in section_name:
        return#process_tbl(tbl,tbl_name)
    elif '评审因素索引表' in section_name:
        return
    elif '施工组织设计' in section_name:
        parent.remove(tbl)
    elif '拟分包项目情况表' in section_name:
        return

#in:tbl_________________________________
#      |       |            |           |
#      | no    |   name     |  content  |
#      |_______|____________|___________|
#      |_______|____________|___________|
#out:[no,name,content]
def get_tbl_header(tbl):
    headers = []
    if not check_tag_is(tbl,'w:tbl'):
        return []
    for tr in tbl:
        if not check_tag_is(tr,'w:tr'):
            continue
        for tc in tr:
            if not check_tag_is(tc,'w:tc'):
                continue
            fulltext = ''
            for p in tc:
                if not check_tag_is(p,'w:p'):
                    continue
                text = get_text(p)
                if text is '':
                    continue
                fulltext += text
            headers.append(fulltext)
        return headers
    return headers

#in:tbl
#       table2-1
#      _________________________________
#      |       |            |           |
#      | no    |   name     |  content  |
#      |_______|____________|___________|
#      |_______|____________|___________|
#out:table2-1    
def get_tbl_name(tbl):
    p = tbl.getprevious()
    while not any(key in get_text(p) for key in ['表','录']):
        print(get_text(p))
        if p is None:# or not check_tag_is(p,'w:p'):
            break
        p = p.getprevious()
    if p is not None:
        return p
    else:
        return None

import datetime

def process_tbl(tbl,tbl_name,hint_words):
    cols = 0
    rows = 0
    r_idx = 0
    ####################################
    for tr in tbl:
        if check_tag_is(tr,'w:tblGrid'):
            cols = len(tr)
            keyword_list = ['' for i in range(cols)]
            keyword_tc_dic = {}
            print(cols)
        elif check_tag_is(tr,'w:tr'):
            #
            class_ = analyze_tr(tr)
            print(class_)
            for case in switch(class_):
                if case('all-black'):
                    idx = 0
                    ###########
                    for tc in tr:
                        if not check_tag_is(tc,'w:tc'):
                            continue
                        width = analyze_tc(tc)
                        text =  get_text(tc)
                        print(width,text)
                        if text is not '':
                            keyword_tc_dic[text] = tc
                        for i in range(width):
                            if text is not '':
                                keyword_list[idx] = text
                            idx += 1
                    ###########
                    if 'all-white' in analyze_tr(tr.getnext()):
                        print(keyword_list)
                        for i in range(len(keyword_list)-1):
                            if i == len(keyword_list) - 2:
                                break
                            if keyword_list[i] == keyword_list[i+1]:
                                keyword_list.remove(keyword_list[i+1])
                        #keyword_list = list(set(keyword_list))
                        #print(keyword_list)
                        if len(keyword_list) is not (len(tr.getnext()) - 1):
                            print('keyword-list makes problem')
                    break
                if case('zebra'):#left-right form
                    idx = 0
                    ###########
                    for tc in tr:
                        if not check_tag_is(tc,'w:tc'):
                            continue
                        if get_text(tc) is '':
                            key_word = get_text(tc.getprevious())
                            fill_in_word = search_word_in_db(tbl_name,key_word,hint_words)
                            if fill_in_word is None:
                                continue
                            if isinstance(fill_in_word,float):
                                fill_in_word = str(int(fill_in_word))
                            elif isinstance(fill_in_word,int):
                                fill_in_word = str(fill_in_word)
                            elif isinstance(fill_in_word,datetime.datetime):
                                continue
                            print(key_word,fill_in_word,hint_words)
                            fill_in_tc(tc,fill_in_word)
                            print('filled')
                            
                    break
                if case('one-black'):#top-down form
                    idx = 0
                    ###########
                    for tc in tr:
                        if not check_tag_is(tc,'w:tc'):
                            continue
                        if check_tag_is(tc,'w:tc') and not check_tag_is(tc.getprevious(),'w:tc'):
                            help_word = get_text(tc)
                            continue
                        key_word = keyword_list[idx]
                        fill_in_word = search_word_in_db(tbl_name,key_word,help_word)
                        if fill_in_word is None:
                            idx += 1
                            continue
                        if isinstance(fill_in_word,float):
                            fill_in_word = str(int(fill_in_word))
                        elif isinstance(fill_in_word,int):
                            fill_in_word = str(fill_in_word)
                        elif isinstance(fill_in_word,datetime.datetime):
                            idx += 1
                            continue
                        print(key_word,fill_in_word,help_word)
                        if help_word in keyword_tc_dic.keys():
                            _tc_ = create_tc(keyword_tc_dic[key_word],fill_in_word)
                            tc.addprevious(_tc_)
                            tr.remove(tc)
                        else:
                            print('__tc__ is None')
                            _tc_ = None
                            fill_in_tc(tc,fill_in_word)                        
                        idx += 1
                    break
                if case('all-white'):
                    c_idx = 0
                    ###########
                    for tc in tr:
                        if not check_tag_is(tc,'w:tc'):
                            continue
                        if r_idx > len(hint_words)-1 or c_idx > len(keyword_list)-1:
                            break
                        key_word = keyword_list[c_idx]
                        fill_in_word = search_word_in_db(tbl_name,key_word,hint_words[r_idx])
                        if fill_in_word is None:
                            c_idx += 1
                            continue
                        if isinstance(fill_in_word,float):
                            fill_in_word = str(int(fill_in_word))
                        elif isinstance(fill_in_word,int):
                            fill_in_word = str(fill_in_word)
                        elif isinstance(fill_in_word,datetime.datetime):
                            c_idx += 1
                            continue
                        print(key_word,fill_in_word,hint_words[r_idx])
                        if key_word in list(keyword_tc_dic.keys()):
                            _tc_ = create_tc(keyword_tc_dic[key_word],fill_in_word)
                            tc.addprevious(_tc_)
                            tr.remove(tc)
                        else:
                            print('__tc__ is None')
                            _tc_ = None
                            fill_in_tc(tc,fill_in_word)
                        c_idx += 1
                    ############
                    r_idx += 1
                    break
            rows += 1
    ####################################
def set_tbl_name(tbl,text):
    p = tbl.getprevious()
    while '表' not in get_text(p):
        if not check_tag_is(p,'w:p') or p is None:
            break
        p = p.getprevious()    
    clear_text(p)
    for r in p:
        if not check_tag_is(r,'w:r'):
            continue
        t = check_tag_exist(r,'w:t')
        if t is not None:
            t.text =  text
            break

def set_title(p,text):
    clear_text(p)
    for r in p:
        if not check_tag_is(r,'w:r'):
            continue
        t = check_tag_exist(r,'w:t')
        if t is not None:
            t.text =  text
            break    
    
def set_keyword_list(tr):
    pass

def set_text(key_word,help_word):
    pass