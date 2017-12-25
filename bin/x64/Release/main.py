# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 23:16:57 2017

@author: Frank
"""

import os

from extract import Extract
from split import Split
from merge import Merge
from fillin import FillIn
from config import db#db_finance,db_company,db_human,db_projects_done,db_projects_being

fullcontent_filename = 'simohua.docx'
input_filepath = os.path.join(os.getcwd(),fullcontent_filename)

extract_filename = 'simohua-extract.docx'
extract_filepath = os.path.join(os.getcwd(),extract_filename)

output_filename = 'output.docx'
output_filepath = os.path.join(os.getcwd(),output_filename)

tmp_dir1 = os.path.join(os.getcwd(),'tmp\\split')
tmp_dir2 = os.path.join(os.getcwd(),'tmp\\split-processed')

from misc import mkdir
    
if __name__ == "__main__":
    #
    extract = Extract(input_filepath)
    extract.process()#output_filepath=extract_filepath)
    db['project_info'].set_db(extract.extract_project_infos())
    #
    mkdir(tmp_dir1)
    split = Split(input_filepath=extract_filepath)
    sections = split.process()
    #
    db['finance'].filtering(need_years=3)
    db['human'].select_people(name_list=['总经理姓名','联系人姓名','项目经理人姓名'])
    db['projects_done'].filtering(project_types=['水利'],need_years=3)
    db['projects_being'].filtering(project_types=['水利'])
    #
    mkdir(tmp_dir2)
    for section in sections:
        fillin = FillIn(os.path.join(tmp_dir1,section+'.docx'))
        fillin.process(os.path.join(tmp_dir2,section+'.docx'))
    #
    merge = Merge(tmpdir=tmp_dir2,section_names=sections)
    merge.process(output_filepath)
    #