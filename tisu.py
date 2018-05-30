# -*- coding: utf-8 -*-
"""
Created on Thu May 24 09:33:50 2018

@author: bin
"""

import pandas as pd
import os
from pandas.tseries.offsets import *
import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime
import time

last2day = time.localtime(time.time()-48*3600)
ts_name = time.strftime('%m%d',last2day)+'.txt'

matplotlib.rcParams['font.sans-serif'] = ['KaiTi']#作图的中文
matplotlib.rcParams['font.serif'] = ['KaiTi']#作图的中文
plt.rcParams['axes.unicode_minus'] = False#正常显示负号

main_file = os.path.join('002-new_data','TISU')
mid_file = '001-mid&old_data'
rst_file = '003-result'
ts_new_path = os.path.join(main_file,ts_name)
ts_old_path = os.path.join(mid_file,'008-tisu.txt')
ts_rst_path = os.path.join(rst_file,'tisu.xlsx')

def read_ts():
    global tisu,tday,tm_day,lm_day,dishi
    rec1 = pd.read_csv(ts_new_path,encoding = 'GBK',dtype=str,sep='\t')
    rec2 = pd.read_csv(ts_old_path,encoding = 'GBK',dtype=str,sep='\t')
    rec1 = rec1[rec1['指标名称']=='当月累计发展量']
    rec2 = rec2[rec2['指标名称']=='当月累计发展量']
    dishi = ['全省','深圳','广州','佛山','东莞',
             '中山','惠州','江门','珠海',
             '汕头','揭阳','潮州','汕尾',
             '湛江','茂名','阳江','云浮',
             '肇庆','梅州','清远','河源','韶关']
    tisu = pd.concat([rec1,rec2])
    tisu['指标值'] = tisu['指标值'].astype(float)
    tday = pd.to_datetime(str(rec1.iloc[-1,0]))
    tm_day = datetime.strftime(tday,'%Y%m%d')
    lm_day = datetime.strftime(tday-DateOffset(months =1),'%Y%m%d')
    
def pdqs():
    qs_txt = ''
    ts_isnull=tisu.groupby(['统计日期','分公司']).sum().unstack()['指标值']
    lable = 0
    for i in ts_isnull.index:
        qs = ts_isnull.loc[i].isnull()
        if qs.any(0):
            queshu = list(ts_isnull.columns[qs])
            temlst = ''.join(['统计日期:',i,'缺数地市:','、'.join(queshu),'\n'])
            qs_txt = qs_txt +temlst
            lable = 1
    if lable == 0:
        qs_txt = ''.join(['统计日期',datetime.strftime(tday,'%Y-%m-%d'),
                          '没有缺数','\n'])
    return qs_txt

def process():
    global tisu,tday,tm_day,lm_day,dishi
    ac_tm = tisu.loc[tisu['统计日期'] == tm_day].groupby('分公司').sum()
    ec_tm = tisu.loc[(tisu['统计日期'] == tm_day) & 
                     (tisu['十六大渠道'] == '电子渠道')].groupby('分公司').sum()
    ec_lm = tisu.loc[(tisu['统计日期'] == lm_day) & 
                     (tisu['十六大渠道'] == '电子渠道')].groupby('分公司').sum()
    rst_rt = ec_tm / ec_lm - 1
    rst_rc = ec_tm / ac_tm
    rst_sum_qs = ec_tm.sum()
    rt_qs = ec_tm.sum() / ec_lm.sum() -1
    rc_qs = ec_tm.sum() / ac_tm.sum()
    rst_qs =pd.DataFrame({'入网量':rst_sum_qs,'环比':rt_qs,'渠道占比':rc_qs})
    rst_qs.index=['全省']
    rst =pd.concat([ec_tm,rst_rt,rst_rc],axis = 1)
    rst.columns = ['入网量','环比','渠道占比']
    rst_addqs = pd.concat([rst,rst_qs])
    rst_addqs = rst_addqs.reindex(index=dishi, columns=['入网量','环比','渠道占比'])
    rst_addqs.fillna(0, inplace=True)
    rst_addqs.to_excel(ts_rst_path)
    return
    

















    
    