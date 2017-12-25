# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 16:20:27 2017

@author: Frank
"""

from lxml import etree
from docx.oxml.ns import qn
import copy

word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
xml_schema = 'http://www.w3.org/XML/1998/namespace'

def check_attrib_exist(element,attrib_name):
    if attrib_name in element.attrib.keys():
        return True
    else:
        return False

def check_attrib_is(element,attrib_name,attrib_value='preserve'):
    if get_attrib_value(element,attrib_name) is attrib_value:
        return True
    else:
        return False

def check_tag_is(element, tag_name):
    if element is None:
        return False
    return element.tag == qn(tag_name)
 
def check_tag_exist(element, tag_name):
    if element is None:
        return None
    for child in element:
        if check_tag_is(child, tag_name):
            return child
    return None

def clear_text(tc):
    for node in tc.iter(tag=etree.Element):
        if node is not None and check_tag_is(node, 'w:t'):
            node.text = ''

def create_p(text=None,fontsize=21):
    #print('.')
    if text is not None:
        p = create_tag('w:p',[['w:rsidR','00486677'],['w:rsidRDefault','00C91861']])
        pPr = create_pPr(fontsize=fontsize)
    else:
        p = create_tag('w:p',[['w:rsidR','00486677'],['w:rsidRDefault','00486677']])
        pPr = create_tag('w:pPr')
        spacing = create_tag('w:spacing',[['w:before','9'],['w:after','0'],['w:line','100'],['w:lineRule','exact']])
        pPr.insert(0,spacing)
        rPr = create_rPr(fontsize=fontsize)
        spacing.addnext(rPr)
    #print('...')
    p.insert(0,pPr)
    #print('....')
    if text is None:
        #r = create_r()
        #pPr.addnext(r)
        pass
    else:
        idx = 1
        for char in text:
            p.insert(idx,create_r(char,fontsize))
            idx += 1
    return p
    #print('.....')

def create_pPr(fontsize=21):
    #print('.')
    pPr = create_tag('w:pPr')
    #print('..')
    tabs = create_tabs()
    #print('...')
    pPr.insert(0,tabs)
    #print('....')
    spacing = create_tag('w:spacing',[['w:after','0'],['w:line','240'],['w:lineRule','auto']])
    #print('.....')
    tabs.addnext(spacing)
    #print('......')
    ind = create_tag('w:ind',[['w:left','452'],['w:right','-20']])
    #print('.......')
    spacing.addnext(ind)
    #print('........')
    rPr = create_rPr(fontsize=fontsize)
    #print('.........')
    ind.addnext(rPr)
    #print('..........')
    return pPr
    
def create_r(text=None,fontsize=21):
    r = create_tag('w:r')
    rPr = create_rPr(fontsize=fontsize)
    r.insert(0,rPr)
    if text is not None:
        t = create_tag('w:t')
        t.text = text
        rPr.addnext(t)
    return r        

def create_rPr(fontsize=21):
    rPr = create_tag('w:rPr')
    rFonts = create_tag('w:rFonts',[['w:ascii','微软雅黑'],['w:eastAsia','微软雅黑'],['w:hAnsi','微软雅黑'],['w:cs','微软雅黑']])
    rPr.insert(0,rFonts)
    sz = create_tag('w:sz',[['w:val',str(fontsize)]])
    rFonts.addnext(sz)
    szCs = create_tag('w:szCs',[['w:val',str(fontsize)]])
    sz.addnext(szCs)
    return rPr

def create_tabs():
    tabs = create_tag('w:tabs')
    tab = create_tag('w:tab',[['w:val','left'],['w:pos','860']])
    tabs.insert(0,tab)
    return tabs
    
def create_tag(tag_name='w:u',attribs = []):
    tag = etree.Element(qn(tag_name))
    for attrib in attribs:
        attrib_name,attrib_value = attrib
        tag.set(qn(attrib_name),attrib_value)
    return tag
    
def del_attribs(element,attrib_names):
    for attrib_name in attrib_names:
        del element.attrib[attrib_name]

def del_tags_by_name(element,tag_names=['w:pgNumType','w:footerReference']):
    if not check_tag_is(element,'w:p') or not check_tag_is(element[0],'w:pPr'):
        return
    for child in element[0]:
        if check_tag_is(child,'w:sectPr'):
            for grandchild in child:
                if grandchild is None:
                    continue
                for tag_name in tag_names:
                    if check_tag_is(grandchild,tag_name):
                        parent = grandchild.getparent()
                        parent.remove(grandchild)
                        grandchild = None
                        break

# if element.text is  网址：地址
# divide it into three,that's,make 2 additional items, one before and one after
# in:  网址：地址
# out: 网址
#       ：
#      地址
def divide_text(r,criteria):
    text = get_text(r)
    if len(text) is not 1:
        previous_letters = text[:text.index(criteria)]
        next_letters = text[text.index(criteria)+1:]
        if previous_letters is not '':
            insert_previous_ritem(r,previous_letters)
        if next_letters is not '':
            insert_next_ritem(r,next_letters)
        r.find(qn('w:t')).text = criteria
        
def find_element_by_tag_name(tree,tagname):
    return tree.find(tagname, tree.nsmap)

def find_index_by_tag_name(element,tagname):
    if element is None:
        return -1
    for index,child in enumerate(element):
        if check_tag_is(child,tagname):
            return index
    return -1

def find_elements_by_text(r,text):
    elements_list = []
    if len(text) < 1:
        return elements_list
    index = 0
    while get_text(r) == text[index] or get_text(r) == '':
        if get_text(r) != '':
            elements_list.append(r)
            index += 1
            if index == len(text):
                break
        r = r.getnext()
    if index < len(text):
        return []
    return elements_list
#def find_child_element_by_tag_name(element,tagname):
#    return element.find('{%s}%s' % (word_schema,tagname))
#def find_element_by_text(tree,text):
#    return tree.xpath('.//a[text()="' + text + '"]')

#def get_attrib_name(tree,btag="xml",stag="space"):
#    return etree.QName(tree.nsmap[btag],stag)

def get_attrib_value(element,attrib_name):
    return element.attrib[attrib_name]#element.get(attrib_name)

def get_attrib_value_word(element,attrib_name):
    return  element.attrib[qn(attrib_name)]

def get_colon_element(element):
    if check_tag_is(element,'w:pPr'):
        return None
    for child in element:
        if check_tag_is(child,'w:t') and child.text is '：':
            print(child.text)
            return child
    return None
        
def get_next_letter_element(element):
    text = get_text(element)
    while  text is "":
        element = element.getnext()
        text = get_text(element)            
    return element

def get_next_string_element(element):
    text = get_text(element)
    # reach to the first letter
    while  text is "" and element.getnext() is not None:
        element = element.getnext()
        text = get_text(element)
    # reach to the last letter
    fulltext = text
    while text is not "" and element.getnext() is not None:
        element = element.getnext()
        text = get_text(element)
        fulltext += text
    #fulltext = fulltext[::-1]
    return element,fulltext

def get_previous_string_element(element):
    text = get_text(element)
    # reach to the first letter
    while  text is "" and element.getprevious() is not None:
        element = element.getprevious()
        text = get_text(element)
    # reach to the last letter
    fulltext = text[::-1]
    while text is not "" and element.getprevious() is not None:
        element = element.getprevious()
        text = get_text(element)
        fulltext = text[::-1]+fulltext#[::-1]
        
    return element,fulltext

def get_preserve_element(element):
    if element is None:
        return None
    if check_tag_is(element,'w:pPr'):
        return None
    for child in element:
        if check_tag_is(child,'w:t') and len(child.attrib) > 0:
            #del child.attrib['{%s}%s' % (xml_schema,'space')]
            return child
    return None

def get_text(element):
    string=''
    for node in element.iter(tag=etree.Element):
        if node is not None and check_tag_is(node, 'w:t'):
            string += node.text if node.text is not None else ''
    return ''.join(string.split())

#parent w:p
#child w:r
#input w:r
#do    w:r
def insert_next_ritem(child,text):
    next_element = copy.deepcopy(child)
    next_element.find(qn('w:t')).text = text#next_element.find('{%s}%s' % (word_schema,'t')).text = text
    parent = child.getparent()
    parent.insert(parent.index(child)+1,next_element)
    return next_element

#parent w:p
#child w:r
#input w:r
#do    w:r
def insert_previous_ritem(child,text):
    previous_element = copy.deepcopy(child)
    previous_element.find(qn('w:t')).text = text#previous_element.find('{%s}%s' % (word_schema,'t')).text = text
    parent = child.getparent()
    parent.insert(parent.index(child),previous_element)
    return previous_element

def keyword_is_previous_string(element,keyword,criteria = '：'):
    if keyword is '':
        return True
    #etree.ElementBase.index()
    match_count = 0        
    while match_count < len(keyword) and element.getprevious() is not None:
        element = element.getprevious()
        text = get_text(element)
        #print(keyword,text)
        if text is '':
            continue
        elif text in keyword or keyword in text:
            match_count += len(text)
        else:
            break
    if match_count == len(keyword):
        return True
    else:
        return False

def keyword_is_next_string(element,keyword):
    #etree.ElementBase.index()
    if keyword is '':
        return True
    match_count = 0
    fulltext = ''
    while match_count < len(keyword) and element.getnext() is not None and len(fulltext) < len(keyword):
        element = element.getnext()
        text = get_text(element)
        if text is '':
            continue
        text = text.replace('）','')
        fulltext += text
        if text in keyword or keyword in text:
            match_count += len(text)
        else:
            break
    if match_count == len(keyword):
        return True
    else:
        return False

def remove_all_childs_before_index(element,index,preserve_indexes=[0]):
    for i,child in enumerate(element):
        if i > index:
            break
        if i in preserve_indexes:
            continue
        parent = child.getparent()
        parent.remove(child)
        child = None    

def remove_all_childs_after_index(element,index):
    for i,child in enumerate(element):
        if i <= index:
            continue
        parent = child.getparent()
        parent.remove(child)
        child = None    

from lxml.etree import strip_tags
def remove_childs_by_tag_names(element,tagnames):
    strip_tags(element,[qn(tagname) for tagname in tagnames])#strip_tags(element,['{%s}%s' % (word_schema,tagname) for tagname in tagnames])
    #for tagname in tagnames:
    #    strip_tags(element,'{%s}%s' % (word_schema,tagnames))
    
def set_attrib(element,attrib_name,attrib_value):
    element.set(attrib_name,attrib_value)
    
def set_attrib_preserve(element):
    if check_tag_is(element,'w:t'):
        element.set(qn('xml:space'),'preserve')#element.set('{%s}%s' % (xml_schema,'space'),'preserve')
        return  
    elif check_tag_is(element,'w:r'):
        for child in element:
            set_attrib_preserve(child)
            #if check_tag_is(child,'t'):
            #    child.set('{%s}%s' % (xml_schema,'space'),'preserve')
    else:
        return

def set_underline(element):
    if check_tag_is(element,'w:rPr') and check_tag_exist(element,'w:u') is None:
        print('making underline')
        eindex = find_index_by_tag_name(element,'w:lang')
        sindex = find_index_by_tag_name(element,'w:szCs')
        if sindex is not -1:
            index = sindex + 1
        elif eindex is not -1:
            index = eindex
        else:
            index = len(element)
        print(index)
        u = etree.Element(qn('w:u'))#u = etree.Element('{%s}%s' % (word_schema,'u'))
        u.set(qn('w:val'),'single')#u.set('{%s}%s' % (word_schema,'val'),'single')
        u.set(qn('w:color'),'000000')#u.set('{%s}%s' % (word_schema,'color'),'000000')
        element.insert(index,u)
    elif check_tag_is(element,'w:r'):
        rPr = check_tag_exist(element,'w:rPr')
        if rPr is not None:
            set_underline(rPr)
    elif check_tag_is(element,'w:p'):
        for child in element:
            set_underline(child)
    else:
        return
        