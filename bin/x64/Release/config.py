# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 16:11:37 2017

@author: Frank
"""
import os
from db import ProjectInfo,Finance,Company,Project,HumanResource

project_info = ProjectInfo()
company = Company(os.path.join(os.getcwd(),'db\\公司.xlsx'))
human = HumanResource(os.path.join(os.getcwd(),'db\\员工.xlsx'))
finance = Finance(os.path.join(os.getcwd(),'db\\财务.xlsx'))
projects_done = Project(os.path.join(os.getcwd(),'db\\往期工程.xlsx'))
projects_being = Project(os.path.join(os.getcwd(),'db\\正在建设项目.xlsx'))

db = {}
db['project_info'] = project_info
db['company'] = company
db['project_members'] = {}
db['human'] = human
db['finance'] = finance
db['projects_done'] = projects_done
db['projects_being'] = projects_being
db['公司信息路径'] = os.path.join(os.getcwd(),'db')
db['人员证件路径'] = os.path.join(os.getcwd(),'img\\证件\\人员')
db['公司证件路径'] = os.path.join(os.getcwd(),'img\\证件\\公司')

PATH = ['公司信息路径','人员证件路径','公司证件路径']

POSITION = [
             '法人',
             '联系人',
             '委托人1',
             '委托人2',
             '项目经理',
             '技术负责人',
             '安全员',
             '材料员',
             '预算员',
             '施工员',
             '质量员',
             '财务人'
             ]

SELECTED = '法人'
#for position in positions:
#    db['project_members'][position] = '陌生人'

INDEX = {
        '投标文件格式':[-1],
        '封面页':['project_info','company'],
        '目录':['no'],
        '投标函及投标函附录':['project_info','company','human'],
        '投标函':['project_info','company','human'],
        '投标函附录':['project_info','company'],
        '承诺书':['company'],
        '法定代表人身份证明':['company','human'],
        '授权委托书':['project_info','company','human'],
        '联合体协议书':['no'],
        '投标保证金':['no'],
        '已标价工程量清单':['no'],
        '施工组织设计':['no'],
        '项目管理机构表':['human'],
        '项目管理机构组成表':['human'],
        '主要人员简历表':['human'],
        '拟分包项目情况表':['no'],
        '资格审查资料':['company','human'],
        '投标人基本情况表':['company','human'],
        '组织机构图':['no'],
        '近3年财务状况表':['finance'],
        '拟投入本项目的流动资金函':['project_info','company'],
        '近3年完成的类似项目情况表':['projects_done'],
        '正在施工的和新承接的项目情况表':['projects_being'],
        '近3年发生的诉讼及仲裁情况表':['no'],
        '资格审查自审表':['company','human','no'],
        '原件的复印件':['no'],
        '其他资料':['no'],
        '其他材料':['no'],
        '资格审查原件登记表':['company'],
        '符合性审查表':['project_info','human','company'],
        }

KEYWORDS = [
            '目录',
            '投标文件格式',
            '招标公告',
            '投标人须知',
            '评标办法',
            '评标方法'
           ]

COVER = [
         '投标文件格式',
         '招标编号',
         '项目名称',
         '投标文件',
         '投标人',
         '盖单位章',
         '法定代表',
         '签字',
         '年月日',
        ]

CONTENT = [
           '评审因素索引表',
           '标段名称',
           '招标文件',
           '投标总报价',
           '项目负责人',
           '技术负责人',
           '投标人名称',
           '招标人名称',
           '授权委托书',
           '身份证号码',
           '委托代理人',
           '法定代表人',
           '法定代表人身份证复印件',
           '代理人身份证复印件',
           '异议函'
           ]

CERTIFICATE = [
               '建造师证',
               '安考A证',
               '安考B证',
               '安考C证',
               '上岗证',
               '会计证',
               '职称证',
               '身份证',
               '学历证',
               '毕业证'
               ]

CERTIFICATE_SIZE = {
                    '建造师证':6.2,
                    '安考A证':6.2,
                    '安考B证':6.2,
                    '安考C证':6.2,
                    '上岗证':6.2,
                    '会计证':6.2,
                    '职称证':6.2,
                    '身份证':2.5,
                    '学历证':6.2,
                    '毕业证':6.2,
                    }

PROJECT_INFO_KEY = [
                    '招标人',
                    '招标人名称',
                    '招标代理机构',
                    '招标编号',
                    '招标范围',
                    '地址',
                    '邮政编号',
                    '邮编',
                    '联系人',
                    '联系电话',
                    '电话',
                    '电子邮件',
                    '网址',
                    '开户单位',
                    '开户银行',
                    '开户账号',
                    '项目名称',
                    '建设地点',
                    '项目基本情况',
                    '项目概况',
                    '合同工期',
                    '计划工期',
                    '计划开工日期',
                    '工期要求',
                    '质量标准',
                    '质量要求',
                    ]

SYNONYMS = [
            ['投标人','投标人名称','本公司名称','名称'],
            ['招标人','招标人名称'],
            ['本人'],
            ['法人','法人代表','法定人','法定代表人','公司代表人','总经理'],
            ['联系人','财务人','财务负责人'],
            ['项目经理','项目负责人','项目经理人'],
            ['委托人','委托代理人'],
            ['评标方法','评标办法'],
            ['办法','方法'],
            ['邮政编号','邮编'],
            ['项目基本情况','项目概况'],
            ['居民身份证编号','身份证号码'],
            ]

DIC = list(INDEX.keys()) + COVER + CONTENT + POSITION
DIC += company.fields + human.fields + finance.fields +  projects_done.fields