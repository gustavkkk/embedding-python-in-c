# -*- coding: utf-8 -*-
"""
Created on Mon Dec 11 10:11:07 2017

@author: Frank

attention: == and is
           == for number
           is for string
           is doesn't work well for number
"""

c2n     = {'零':0,
           '一':1,
           '二':2,
           '三':3,
           '四':4,
           '五':5,
           '六':6,
           '七':7,
           '八':8,
           '九':9,
           '十':10,
           '百':100,
           '千':1000,
           '万':10000,
           '亿':100000000,
           }

def n2c(number):
    for key,value in c2n.items():
        if value == number:
            return key
    return ''

def chn2int(chinese):
    if not all(char in c2n.keys() for char in chinese):
        return -10000
    sum = 0
    #for i in range(len(chinese)-1,-1,-1):chinese[::-1]
    i = len(chinese) - 1
    while i > -1:
        cur_ = c2n[chinese[i]]
        print(i,chinese[i],cur_)
        if i == 0:
            sum += cur_
            return sum
        if cur_ < 10:
            sum += cur_
            i -= 1
            continue
        elif cur_ < 10000:
            i -= 1
            mul_ = c2n[chinese[i]]
            sum += mul_ * cur_
            i -= 1
            continue
        elif cur_ == 10000:
            if not '亿' in chinese:
                print(sum)
                print(chinese[:i-1])
                sum += chn2int(chinese[:i-1]) * 10000
                return sum
            else:
                s = chinese.index('亿')
                sum += chn2int(chinese[s:i-1]) * 10000
                i = s
        elif cur_ > 10000:
            sum += chn2int(chinese[:i-1]) * 10000
            return sum            
    return sum
            
def int2chn(int_):
    chinese = ''
    if int_ < 10000:
        unit = 1
        while int_ > 0:
            cur_ = int_ % 10
            print(int_,cur_)
            int_ = int(int_/10)
            if cur_ == 0:
                unit *= 10
                continue
            elif unit == 1:
                chinese = n2c(cur_)
            else:
                chinese = n2c(cur_) + n2c(unit) + chinese
            unit *= 10
        return chinese
    elif int_ >= 10000 and int_ < 100000000:
        chinese = int2chn(int_ % 10000)
        chinese = int2chn(int(int_ / 10000)) + '万' + chinese
        return chinese
    else:
        chinese = int2chn(int_ % 100000000)
        chinese = int2chn(int(int_ / 100000000)) + '亿' + chinese
    return chinese

def get_words_between_brackets(text,opening=['（'],closing=['）']):
    words_list = []
    i = 0
    while i < len(text):
        if text[i] in opening:
            s = i
            while i < len(text):
                if text[i] in closing:
                    words_list.append([text[s+1:i],s+1,i])
                    break
                elif text[i] in opening:
                    s = i
                i += 1
        i += 1
    return words_list

def recommend(title):
    next_strings = []
    for item in get_words_between_brackets(title):
        string,s,e = item
        if all(char in list(c2n.keys()) for char in string):
            next_strings.append(title[:s]+n2c(c2n[string]+1)+title[e:])
    return next_strings
''''''
from .tag import check_tag_is,get_text
import copy
#from tagprocessing import del_tags      
def extract_pages(body,firstindex,lastindex):
    for para_index,child in enumerate(body):
        if para_index < firstindex or para_index >lastindex or check_tag_is(child,'w:sectPr'):
            #del_tags(child)
            parent = child.getparent()
            parent.remove(child)
            child = None

def split_page_by_brackets(xtree,section_name):
    section_breaks = []
    #print(section_name)
    section_breaks.append([section_name,0])
    for i,p in enumerate(xtree[0]):
        if not check_tag_is(p,'w:p'):
            continue
        title = get_text(p)
        for item in get_words_between_brackets(title):
            string,s,e = item
            if all(char in list(c2n.keys()) for char in string):
                section_breaks.append([title,i])
    #print(section_breaks)
    sections = []
    for index,section in enumerate(section_breaks):
        section_name,first_para_index = section
        xtree_ = copy.deepcopy(xtree)
        xbody = xtree_[0]
        if index == len(section_breaks) - 1:
            last_para_index = len(xbody) - 1
        else:
            last_para_index = section_breaks[index+1][1] - 1
        extract_pages(xbody,first_para_index,last_para_index)
        sections.append([xtree_,section_name])
    return sections
''''''                
if __name__ == "__main__":
    #print(chn2int('八百九十九万一千三百二十九'))
    #print(int2chn(12220007))#2212222341))
    print(recommend('授权委托书（一）（格式）'))