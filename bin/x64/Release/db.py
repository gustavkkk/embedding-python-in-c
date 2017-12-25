# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 09:21:40 2017

@author: Frank

this file contains excel-handling-functions.
primitive now, needs more progress.
modification of db, synonym-based getdata function,...
"""

import pandas as pd
from pandas import ExcelFile,ExcelWriter
import os
import numpy as np
import datetime

class ProjectInfo:
    def __init__(self):
        self.dic = {}

    def set_db(self,dic):
        self.dic = dic
        
    def add_data(self,field,value):
        self.dic[field] = value
        
    def get_data(self,field):
        if field in self.dic.keys():
            return self.dic[field]
        else:
            return None
        
class EXCEL:
    def __init__(self,xls_filepath):
        self.filepath = xls_filepath
        self.xls_reader = ExcelFile(xls_filepath)
        self.sheet_names = self.xls_reader.sheet_names
        if len(self.sheet_names) == 1:
            self.select_sheet(self.sheet_names[0])
        self.time = datetime.datetime.now()

    def add(self):
        pass
    
    @property
    def data(self):
        return self._data
    
    def select_sheet(self,sheet_name):
        self._data = self.xls_reader.parse(sheet_name)#self._data = pd.read_excel(xls_filepath)

    def merge_sheet(self):
        sheets = []
        for sheet_name in self.sheet_names:
            sheet = self.xls_reader.parse(sheet_name)
            sheets.append(sheet)
        self._data = pd.concat(sheets)
        
    def save(self,xls_filepath,sheet_name='Sheet5'):
        self.xls_reader.close()
        self.xls_writer = ExcelWriter(xls_filepath)
        self._data.to_excel(self.xls_writer,sheet_name)
        self.xls_writer.save()
    
class Finance(EXCEL):
    def __init__(self,xls_filepath,need_years=3):
        super(Finance, self).__init__(xls_filepath)
        self.merge_sheet()
        self.reset_row_col()
        self.filtering(need_years=need_years)
       
    def add(self):
        pass
    
    def get_cell(self,when='2014年',what='净资产'):
        if not when in self.years or not what in self.rows:
            print("no such an item")
            return None
        return self.selected[self.dic_row[what]][self.dic_col[when]]

    def get_row(self,when='2014年'):
        if not when in self.years:
            return "no such a column"
        return self._data[when].values

    def reset_row_col(self,row_identifier=r'名称'):
        #print(self.data.values)
        self.rows = list(self._data[row_identifier].values)#list(self._data.index)
        self.cols = list(self._data.columns)
        if len(self.cols) > 1:
            self.cols.pop(0)
            
    def filtering(self,need_years=3):
        # select years for reporting
        if need_years >= len(self.years):
            cols = self.years[-need_years:]
        else:
            cols = self.years
        #
        self.selected = np.array(self.data[cols])
        # create dictionary
        self.dic_row = {}
        for row_index,name in enumerate(self.rows):
            self.dic_row[name] = row_index
        self.dic_col = {}
        for col_index,name in enumerate(cols):
            self.dic_col[name] = col_index

    @property    
    def fields(self):
        return self.rows
    
    @property    
    def years(self):
        return self.cols

class Company(EXCEL):
    
    def __init__(self,xls_filepath):
        super(Company, self).__init__(xls_filepath)
        self.select_sheet(self.sheet_names[0])
        self._fields = list(self._data['FIELD'].values)
        self._values = list(self._data['VALUE'].values)
        
    def get_data(self,what='名称'):
        if what not in self.fields:
            print("no such an item")
            return None
        return self.values[self.fields.index(what)]

    @property    
    def fields(self):
        return self._fields
    
    @property
    def values(self):
        return self._values

class Project(EXCEL):
    def __init__(self,xls_filepath):
        super(Project, self).__init__(xls_filepath)
        self.select_sheet(self.sheet_names[0])
        self.cols = list(self.data.columns)
        
    def filtering(self,project_types=None,need_years=None):
        self._data.sort_values(by='年度')#?
        if project_types is None:
            self.filtered = np.array(self.data)
            return self.data
        filtered = self.data[self.data['类型'].isin(project_types)]
        if need_years is not None:
            years = [self.time.year - i - 1 for i in range(0,need_years)]
            filtered = filtered[filtered['年度'].isin(years)]
        self.filtered = np.array(filtered)
        return filtered
    
    def get_data(self,index=0,what='合同名称'):
        if what not in self.cols:
            return None
        return self.filtered[index][self.cols.index(what)]
    
    @property    
    def fields(self):
        return self.cols

class NotFoundException(BaseException):
    def __init__(self):
        print('not found')
        
class HumanResource(EXCEL):
    def __init__(self,xls_filepath):
        super(HumanResource, self).__init__(xls_filepath)
        self.select_sheet(self.sheet_names[0])
        self.cols = list(self.data.columns)
        self.filtered = None
        
    def select_people(self,name_list=['王养','联系人','秦仁和']):
        self.filtered = self.data[self.data['姓名'].isin(name_list)]
        #print(self.filtered['姓名'])
        pass

    def filtering_by_position(self,positions=['经理']):
        #filtered = self.data[self.data['职务'].isin([position])]
        #filtered = self.data.query('["'+position+'"] in 职务')
        #print(self.data['职务'])
        indexlist = []
        for index,item in enumerate(self.data['职务']):
            if any(position in item for position in positions):
                indexlist.append(index)
        filtered = self.data[self.data.index.isin(indexlist)]
        return np.array(filtered)[:,self.cols.index('姓名')]
    
    def get_data(self,position=None,name=None,what='年龄'):
        if position is None and name is None:
            print('input condition')
            return None
        elif self.filtered is None:
            print('not selected')
            return None
        elif position is not None:
            filtered = self.filtered[self.data['职务'].isin([position])]
        else:
            filtered = self.filtered[self.data['姓名'].isin([name])]
        #print(filtered)
        filtered = np.array(filtered)
        if filtered.shape[0] > 1:
            #raise(NotFoundException)
            print('not one with same condition')
            return None
        elif filtered.shape[0] == 0:
            #raise(NotFoundException)
            print('not found')
            return None
        else:
            return filtered[0][self.cols.index(what)]
        
    @property    
    def fields(self):
        return self.cols

def similarity(word1,word2):
    return 0.0

def find_synonym(src_word='',target_list=[]):
    for dest_word in target_list:
        if similarity(src_word,dest_word) > 0.9:
            return dest_word
    return None        
    
fullpath = os.path.join(os.getcwd(),'db\\财务.xlsx')

def testfinance():
    finance = Finance(os.path.join(os.getcwd(),'db\\财务.xlsx'))
    print(finance.get_cell('2014年','净资产'))
    print(finance.years)
    print(finance.types)

def testcompany():
    company = Company(os.path.join(os.getcwd(),'db\\公司.xlsx'))
    print(company.get_data('成立时间'))
    #print(company.fields)
    #print(company.values)

def testproject():
    prj = Project(os.path.join(os.getcwd(),'db\\往期工程.xlsx'))
    print(prj.filtering(project_types=['水利']))
    print(prj.get_data())
    print(prj.fields)

def testhuman():
    human = HumanResource(os.path.join(os.getcwd(),'db\\员工.xlsx'))
    #human.select_people()
    #print(human.get_data())
    print(human.filtering_by_position())
    
if __name__ == "__main__":
    testhuman()