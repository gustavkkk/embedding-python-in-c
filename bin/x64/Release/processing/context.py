# -*- coding: utf-8 -*-
"""
Created on Mon Nov 27 10:59:41 2017

@author: Frank
"""

from .tag import get_text,divide_text,check_tag_is,check_tag_exist,check_attrib_exist
from .tag import insert_next_ritem,insert_previous_ritem
from .tag import get_next_string_element,get_preserve_element
from .tag import get_next_letter_element,get_previous_string_element
from .tag import keyword_is_previous_string,keyword_is_next_string
from .tag import remove_all_childs_before_index,remove_all_childs_after_index
#from .tag import find_child_element_by_tag_name
from docx.oxml.ns import qn
import jieba
#http://blog.csdn.net/qq_26376175/article/details/69680992
jieba.set_dictionary("./dict.txt")  
jieba.initialize()
#from jieba import analyse  
#jieba.analyse.set_idf_path("./idf.txt")
for word in ['法定代表人','单位性质','年月日','身份证号码','委托代理人','经营期限','投标人名称','项目名称','成立时间']:
    jieba.add_word(word)
from config import PROJECT_INFO_KEY
for word in PROJECT_INFO_KEY:
    jieba.add_word(word)    
#
def correct_margin(p):
    # correct left and right margin
    if check_tag_is(p[0],'w:pPr'):
        tag_ind = check_tag_exist(p[0],'w:ind')
        if tag_ind is not None:
            #print(tag_ind.attrib.keys())
            left_attrib = qn('w:left')
            right_attrib = qn('w:right')
            if check_attrib_exist(tag_ind,left_attrib) and check_attrib_exist(tag_ind,right_attrib):
                #left = tag_ind.attrib[left_attrib]
                right = tag_ind.attrib[right_attrib]
                #print('correct_margin',left,right)
                #tag_ind.attrib['{%s}%s' % (word_schema,'left')] = 
                if int(right) > 100:
                    tag_ind.attrib[right_attrib] = str(-20)

# 'while' doesn't work
# change it into 'for'
def eliminate_unclear_clusters(p):
    r = p[0]
    while r is not None:
        text = get_text(r)
        for key in ['：','（','）','；',':','(',')']:
            if key in text:
                divide_text(r,key)
                break
        r = r.getnext()

# 'while' doesn't work
# change it into 'for'
def eliminate_unnecessary_preserve_elements(p):
    r = p[0]
    while r is not None:
        if get_preserve_element(r) is not None:
            next_r = r.getnext()
            while get_preserve_element(next_r) is not None:
                p.remove(next_r)
                next_r = r.getnext()
        r = r.getnext()    
# 'while' doesn't work
# change it into 'for'
def eliminate_noproof_elements(p):
    for r in p:
        if not check_tag_is(r,'w:r'):
            continue
        for t in r:
            if check_tag_is(t,'w:rPr') and check_tag_exist(t,'w:noProof') is not None:
                #print('noProof exists')
                p.remove(r)
        
false_keywords = ['声明','日期']

def merge_two_dicts(x, y):
    z = x.copy()   # start with x's keys and values
    z.update(y)    # modifies z with y's keys and values & returns None
    return z

#
# 2.1 招标编号：EVDW50154
# 2.3 招标范围：工程量清单及施工设计图纸内的全部内容
# 招标人名称：巴东县水土保持局
# 
#
def extract_info_from_paras(body):
    whole = {}
    for i,p in enumerate(body):
        if not check_tag_is(p,'w:p'):
            continue
        text = get_text(p)
        if text == '':
            continue
        individual = extract_info_by_colon(text)
        for key in list(individual.keys()):
            if key not in list(whole.keys()):
                whole[key] = individual[key]
    return whole

from config import PROJECT_INFO_KEY
def extract_info_by_colon(p_str):
    seglist = segment(p_str)
    info = {}
    for i,seg in enumerate(seglist):
        if any(mark in seg for mark in ["：",":"]) and i > 0:
            key = seglist[i-1]
            if not key in PROJECT_INFO_KEY:
                continue
            value = ''
            for j in range(i+1,len(seglist)):
                word = seglist[j]
                if word in PROJECT_INFO_KEY:
                    break
                if any(condition in key for condition in ['邮件','网址','邮编','邮政编号','电话','账号']) and ord(word[0]) >300:
                    break
                if any(condition in key for condition in ['日期','工期']):
                    if any(mark in word for mark in ['：',':',',','，','.','。','(',')','（','）','、']):
                        break
                elif any(condition in key for condition in ['情况','概况','范围']):
                    if any(mark in word for mark in ['：',':']):
                        break
                elif any(condition in key for condition in ['网址']):
                    if any(mark in word for mark in ['。','(',')','（','）','，','、']):
                        break
                else:
                    if any(mark in word for mark in ['。','(',')','（','）','、']):#['：',':',',','，','.','。','(',')','（','）']):
                        break                    
                value += word
            if value is not '':
                info[key] = value
    return info
                
def find_keywords(p_str):
    # segmenting
    #print('.')
    seglist = segment(p_str)
    #log(seglist)
    print(seglist)
    # processing colon
    item_list, remained = find_keywords_by_colon(seglist)
    #log(item_list)
    # processing parenthesis
    #print('..')
    item_list = find_keywords_by_parenthesis(item_list,remained)
    #print('...')
    # fill in
    #print(item_list)
    return item_list

def find_keywords_by_colon(seglist):
    # ：__(), ：
    # ex: 投标人：___（盖单位章）
    #     身份证号码：
    #
    item_list = []
    remained = copy.copy(seglist)
    for i,seg in enumerate(seglist):
        if any(mark in seg for mark in ["：",":"]) and i > 0:
            key_word = seglist[i-1]
            if any(mark in key_word for mark in ['）',')']):
                continue
            if any(key_word in key for key in false_keywords):
                continue
            remained.remove(key_word)
            remained.remove('：')
            if (i+1) < len(seglist) and any(mark in seglist[i+1] for mark in ['（','(']):
                help_word = seglist[i+2]
                for word in ['（','(',help_word,'）',')']:
                    if word in remained:
                        remained.remove(word)
            else:
                help_word = ""
            item_list.append([key_word,help_word,True])
    return item_list,remained

def find_keywords_by_parenthesis(item_list,remained):
    # __()
    # ex: 本人__（姓名）
    #     __(项目名称)   
    for i,seg in enumerate(remained):
        #log(remained)
        if any(mark in seg  for mark in ['（','(']) and (i+2) < len(remained):
            #print(remained[i+1])
            if i == 0 or  any(mark in remained[i-1]  for mark in ['）',')']):
                item_list.append([remained[i+1],"",False])
                #for word in ['（',remained[i+1],'）']:
                #    remained.remove(word)
            else:
                item_list.append([remained[i+1],remained[i-1],False])
                #for word in [remained[i-1],'（',remained[i+1],'）']:
                #    remained.remove(word)
    return item_list

import copy
import datetime
#from .log import log
date = datetime.datetime.today()#datetime.datetime.now()#

def fill_in_paragraph(section_name,p,item_list=None):
    if item_list is None:
        item_list = find_keywords(get_text(p))
    # filling in
    for item in item_list:
        #print(item)
        [key_word,help_word,iscolon] = item
        fillinword = search_word_in_db(section_name,key_word,help_word)
        if fillinword is None:
            #print(item)
            return
        print(fillinword)
        if isinstance(fillinword,float):
            fillinword = str(int(fillinword))
        elif isinstance(fillinword,int):
            fillinword = str(fillinword)
        elif isinstance(fillinword,datetime.datetime):
            #print('setting date.')
            set_date2(p,[str(info) for info in [fillinword.year,fillinword.month,fillinword.day]])
            print(key_word,":",fillinword,help_word)
            continue
        if iscolon:
            #print(key_word,":",fillinword,help_word)
            set_text_by_colon(p,key_word,fillinword,help_word)
        else:
            #print(help_word,fillinword,"("+key_word+')')
            set_text_by_parenthesis(p,key_word,fillinword,help_word)
    #
    #print('filling in date')
    if '年月日' in get_text(p):
        #print('setting date..')
        set_date2(p,[str(info) for info in [date.year,date.month,date.day]])

# element:w:p
# in: w:p,'：',text
# out:
def find_blank_element_index_by_text(p_element,text,criteria='：'):
    isfound_criteria = False
    for i,r_element in enumerate(p_element):
        if check_tag_is(r_element,'w:r'):
            etext = get_text(r_element)
            if etext is '':
                continue
            if not isfound_criteria:
                if criteria in etext:
                    isfound_criteria = True
                continue
            #print(i,text,etext)
            if etext in text:
                return i-1
    return i
    
def preprocess(p):
    # split all string chuncks
    #print('.')
    split_all_clusters(p)
    # split tab chuncks
    #split_tab_clusters(p)
    # eliminate preserve chuncks
    #eliminate_unnecessary_preserve_elements(p)
    # one-keyword,one-paragraph principle
    #elminate noproof
    #print('..')
    eliminate_noproof_elements(p)
    #print('...')
    split_paragraph(p)

from config import INDEX,db#db_finance,db_company,db_human,db_projects_done,db_projects_being
import config as cfg
from misc import switch

def segment(text):
    seglist = jieba.cut(text, cut_all=False)
    return list(seglist)

def search_word_in_db(section_name,keyword,help_word=None):
    result = None
    if section_name not in list(INDEX.keys()):
        for key in list(INDEX.keys()):
            if key in section_name:
                break
        if key not in section_name:
             return None
        else:
             section_name = key
    for db_name in INDEX[section_name]:
        if db_name in db.keys():
            for case in switch(db_name):
                if case('project_info'):
                    if any(condition in keyword for condition in ['招标人','招标编号','项目名称']):
                        result = db[db_name].get_data(keyword)
                    break
                if case('company'):
                    if keyword in ['投标人','投标人名称','本公司名称']:
                        keyword = '名称'
                    result = db[db_name].get_data(what=keyword)
                    break
                if case('human'):
                    #
                    #print(keyword)
                    if any(key in ['法人','法定代表人'] for key in [keyword,help_word]):
                        cfg.SELECTED = '法人'
                    elif any(key in ['委托人','委托代理人','代理人'] for key in [keyword,help_word]):
                        if '委托' in cfg.SELECTED:
                            cfg.SELECTED = '委托人2'
                        else:
                            cfg.SELECTED = '委托人1'
                    if not keyword in db[db_name].fields:
                        break
                    position = None
                    # 依照该项的标题，小建议
                    if '法定代表人' in section_name:
                        position = '法人'
                    elif '授权委托书' in section_name:
                        #position = '法人' or '委托人'
                        if help_word is not None and any(key in help_word for key in ['本人','我']):
                            position = '法人'
                            cfg.SELECTED = '法人'
                        elif help_word is not None and '委托' in help_word:
                            position = '委托人1'
                        if cfg.SELECTED is not None:
                            print(cfg.SELECTED)
                            position = cfg.SELECTED
                    elif help_word is not None and help_word in list(db['project_members'].keys()):
                        position =  help_word
                    ###
                    if position is not None:
                        name = db['project_members'][position]
                        result = db[db_name].get_data(name=name,what=keyword)
                    else:
                        result = None
                    #print(name,keyword,result)
                    break
                if case('finance'):
                    result = db[db_name].get_cell(when='2015年',what=keyword)
                    break
                if case('projects_done'):
                    result = db[db_name].get_data(index=0,what=keyword)
                    break       
                if case('projects_being'):
                    result = db[db_name].get_data(index=0,what=keyword)
                    break
            if result is not None:
                return result
    return result
 
date_index = {'年':0,'月':1,'日':2}        
def set_date(p,date_infos):
    r = p[0]
    while r is not None:
        preserve_element = get_preserve_element(r)
        if preserve_element is not None:
            r = get_next_letter_element(r)
            preserve_element.text = date_infos[date_index[get_text(r)[0]]]
        r = r.getnext()

from .tag import find_elements_by_text,set_attrib_preserve

def set_date2(p,date_infos):
    r = p[0]
    text = '年月日'
    while r is not None:
        #print('...')
        element_list = find_elements_by_text(r,text)
        #print('...')
        for index,element in enumerate(element_list):
            previous_r = element.getprevious()
            info = date_infos[date_index[text[index]]]
            #print(info)
            while check_tag_exist(previous_r,'w:t') is None:
                if previous_r is None:# or get_text(previous_r) is not '':
                    break
                p.remove(previous_r)
                previous_r = element.getprevious()
            if previous_r is None:
                previous_r = insert_previous_ritem(element,info)
            elif get_text(previous_r) is not '':
                previous_r = insert_previous_ritem(element,info)
            text_element = check_tag_exist(previous_r,'w:t')
            set_attrib_preserve(text_element)
            remove_childs_by_tag_names(previous_r,['w:tab'])
            text_element.text = date_infos[date_index[text[index]]]
            r = element
        r = r.getnext()
#
#__(招标人名称)
#巴东县水土保持局(招标人名称)
def set_text_by_underline(p,key_word,fillinword,help_word=''):
    r = p[0]
    while r is not None:
        preserve_element = get_preserve_element(r)
        if preserve_element is not None:
            r,next_str = get_next_string_element(r)
            if help_word is not '':
                r_,previous_str = get_previous_string_element(r)
                print(help_word,previous_str)
            else:
                previous_str = help_word
            if key_word in next_str and help_word in previous_str:
                preserve_element.text = fillinword
        if r is None:
            break
        else:
            r =  r.getnext()

from .tag import set_underline

def set_text_by_parenthesis(p,key_word,fillinword,help_word=''):
    r = p[0]
    while r is not None:
        if '（' in get_text(r) and \
            keyword_is_next_string(r,key_word) and\
            keyword_is_previous_string(r,help_word):
            #autolog('gotcha')
            print(help_word,fillinword,'('+key_word+')')
            previous_r = r.getprevious()
            # find r with t tag
            '''
            while check_tag_exist(previous_r,'w:t') is None:
                if previous_r is None:
                    break
                p.remove(previous_r)
                previous_r = r.getprevious()
            '''
            # find r with preserve
            while get_preserve_element(previous_r) is None:
                if previous_r is None or get_text(previous_r) is not '':
                    break
                p.remove(previous_r)
                previous_r = r.getprevious()
            if previous_r is None:
                previous_r = insert_previous_ritem(r,fillinword)
            elif get_text(previous_r) is not '':
                previous_r = insert_previous_ritem(r,fillinword)
            elif check_tag_exist(previous_r,'w:t'):
                previous_r.find(qn('w:t')).text = fillinword
            else:
                p.remove(previous_r)
                previous_r = insert_previous_ritem(r,fillinword)
            set_attrib_preserve(previous_r)
            set_underline(previous_r)
        r =  r.getnext()
#
#投标人：___(盖单位章)
#法人名称：___(签字)
#地址：
#网址：
            
def set_text_by_colon(p,key_word,fillinword,help_word=''):
    r = p[0]
    while r is not None:
        if '：' in get_text(r) and keyword_is_previous_string(r,key_word):
            next_element = r.getnext()
            if next_element is None:
                next_element = insert_next_ritem(r,fillinword)
            elif get_text(next_element) is not '':
                next_element = insert_next_ritem(r,fillinword)
            elif check_tag_exist(next_element,'w:t'):                    
                next_element.find(qn('w:t')).text = fillinword
            else:
                p.remove(next_element)
                next_element = insert_next_ritem(r,fillinword)                
            set_underline(next_element)
        r =  r.getnext()
        
def split_all_clusters(p):
    r = p[0]
    while r is not None:
        text = get_text(r)
        if len(text) > 1:
            for i,char in enumerate(text):
                #print(text,i,char)
                if i is 0:
                    r.find(qn('w:t')).text = char
                elif i+1 is len(text):
                    break
                insert_next_ritem(r,text[i+1])
                r = r.getnext()
        r = r.getnext()

from .tag import remove_childs_by_tag_names,find_index_by_tag_name

def split_tab_clusters(p):
    r = p[0]
    while r is not None:
        text = get_text(r)
        if len(text) > 0 and check_tag_exist(r,'w:tab') is not None:
            if find_index_by_tag_name(r,'w:t') <\
               find_index_by_tag_name(r,'w:tab'):
                inserted = insert_next_ritem(r,' ')
            else:
                inserted = insert_previous_ritem(r,' ')
            #r.remove(check_tag_exist(r,'tab'))
            #inserted.remove(check_tag_exist(inserted,'tab'))
            remove_childs_by_tag_names(r,['w:tab'])
            remove_childs_by_tag_names(inserted,['w:tab'])
        r = r.getnext()
    
def split_paragraph(p):
    #skip sentence
    #print('.')
    text = get_text(p)
    if "。" in text:
        return
    #print('..')
    #segmenting
    seglist = segment(text)
    #print(seglist)
    #print('...')
    #processing colons
    item_list,remained = find_keywords_by_colon(seglist)
    ###
    #print('....')
    if len(item_list) <= 1:
        return
    ###
    correct_margin(p)
    #print(item_list)
    parent =p.getparent()
    ##
    count = 1
    while count < len(item_list):
        next_p = copy.deepcopy(p)
        next_keyword = item_list[count][0]
        index = find_blank_element_index_by_text(p,next_keyword)
        #print(index,next_keyword)
        remove_all_childs_after_index(p,index)
        remove_all_childs_before_index(next_p,index)   
        ##
        parent.insert(parent.index(p)+1,next_p)
        #
        p = next_p
        count += 1