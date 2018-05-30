# -*- coding: utf-8 -*-
"""
Created on Thu Mar  8 14:10:34 2018

@author: bin
"""

import numpy as np
import pandas as pd
from pandas import Series, DataFrame
from pandas.tseries.offsets import *
import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime
import time
import os
matplotlib.rcParams['font.sans-serif'] = ['KaiTi']#作图的中文
matplotlib.rcParams['font.serif'] = ['KaiTi']#作图的中文
plt.rcParams['axes.unicode_minus'] = False#正常显示负号
last2day = time.localtime(time.time()-48*3600)
itv_name = time.strftime('%m%d',last2day)+'.txt'

main_file = os.path.join('002-new_data','IPTV')
mid_file = '001-mid&old_data'
rst_file = '003-result'
itv_new_path = os.path.join(main_file,itv_name)
itv_old_path = os.path.join(mid_file,'007-iptv.txt')
zddyb_path =  os.path.join(mid_file,'089-zddyb.csv')
itv_rst_path = os.path.join(rst_file,'iptv.xlsx')

#读入文件
def read_itv():
    global iptv,tday,tday_day,tm,lm,dishi
    rec1 = pd.read_csv(itv_new_path,encoding='GBK',dtype=str,sep='\t')
    rec2 = pd.read_csv(itv_old_path,encoding='GBK',dtype=str,sep='\t')
    itv_mapping = pd.read_csv(zddyb_path,encoding = 'GBK',dtype=str)
    tday = pd.to_datetime(str(rec1['统计日期'].iloc[-1]))
    if tday.is_month_end:
        tday_day = 31
    else:
        tday_day = tday.day
    tm = datetime.strftime(tday,'%Y-%m')
    lm = datetime.strftime(tday-DateOffset(months =1),'%Y-%m')
    dishi = ['深圳','广州','佛山','东莞',
             '中山','惠州','江门','珠海',
             '汕头','揭阳','潮州','汕尾',
             '湛江','茂名','阳江','云浮',
             '肇庆','梅州','清远','河源','韶关']
    #开始处理
    iptv = pd.concat([rec1,rec2])
    iptv = pd.merge(iptv,itv_mapping,how='left')
    iptv['ITV新增用户数']=iptv['ITV新增用户数'].apply(lambda x:x.replace(',',''))
    iptv['ITV新增用户数']=iptv['ITV新增用户数'].astype(float)
    iptv['统计日期']=pd.to_datetime(iptv['统计日期'])
    iptv = iptv.set_index('统计日期')

#判断是否缺数
def pdqs():
    qs_txt = ''
    itv_isnull=iptv.groupby(['统计日期','地市']).sum().unstack()['ITV新增用户数']
    lable = 0
    for i in itv_isnull.index:
        qs = itv_isnull.loc[i].isnull()
        if qs.any(0):
            queshu = list(itv_isnull.columns[qs])
            temlst = ''.join(['统计日期:',i.strftime('%Y%m%d'),'缺数地市:','、'.join(queshu),'\n'])
            qs_txt = qs_txt +temlst
            lable = 1
    if lable == 0:
        qs_txt = ''.join(['统计日期',datetime.strftime(tday,'%Y-%m-%d'),'没有缺数','\n'])
    return qs_txt


#数据筛选
def process():
    
    ac = iptv[iptv.index.day<=tday_day]
    ec = ac[(ac['16大渠道']=='电子渠道')]
    ec_last = ec.loc[lm].groupby('地市').sum().copy()
    ac_now = ac.loc[tm].groupby('地市').sum().copy()
    ec_now = ec.loc[tm].groupby('地市').sum().copy()
    
    fig, axes = plt.subplots(2, 2)
    idx = pd.IndexSlice
    pct_ec = ec.groupby(by = ['地市',ec.index]).sum().pct_change().loc[idx[:,tday],:].unstack()
    #看各地市阵地的变化情况
    zd_ds_tm = ec.loc[tm].groupby(['地市','阵地']).sum()
    zd_ds_lm = ec.loc[lm].groupby(['地市','阵地']).sum()
    zd_ds_chge = zd_ds_tm.sub(zd_ds_lm,fill_value = 0)#每个地市每个阵地的变化值
    rst_zd_chge = zd_ds_chge.groupby("地市")['ITV新增用户数'].agg(['idxmax','idxmin'])#每个地市增长最大及最小的阵地
    #结果表
    rst_rt = ec_now/ec_last-1#各地市环比
    rst_rt_qs =ec_now.sum()/ec_last.sum()-1#全省环比
    rst_rc = ec_now/ac_now#各地市渠道占比
    rst_rc_qs =ec_now.sum()/ac_now.sum()#全省渠道占比
    rst_sum=ec_now
    rst_sum_qs = ec_now.sum()
    rst_qs =DataFrame({'入网量':rst_sum_qs,'环比':rst_rt_qs,'渠道占比':rst_rc_qs},columns=['入网量','环比','渠道占比'])
    rst_qs.index=['全省']
    rst =pd.concat([rst_sum,rst_rt,rst_rc,rst_zd_chge],axis = 1)
    col = ['入网量','环比','渠道占比','增长最快','增长最慢']
    rst.columns = col
    rst=rst.reindex(dishi)
    rst_zd = ec.groupby('阵地').resample('M').sum()['ITV新增用户数'].unstack().sort_index()
    rst_zd.columns=rst_zd.columns.to_period('M')
    
    
    #作图中间表
    plot_zrw = ec.resample('M').sum()['ITV新增用户数']
    plot_zrw.index=plot_zrw.index.to_period('M')
    plot_ft = ec[tm].resample('D').sum()['ITV新增用户数']
    plot_ft.index = plot_ft.index.map(lambda x:str(x.day)+'D')
    plot_ds =pd.concat([ec_last,ec_now],axis=1)
    plot_ds.columns = [lm,tm]
    plot_ds.sort_values(by = plot_ds.columns[-1] ,inplace = True,ascending=False)
    #作图
    pct_ec=pct_ec.reindex(plot_ds.index)
    pct_ec.plot(kind = 'bar',ax=axes[0,0])
    plot_ds[:5].plot(kind = 'bar',title = '入网量前五地市情况',ax = axes[0,1])
    plot_ft.plot(kind = 'bar',title = '当月各日入网量',ax = axes[1,0])
    rst_zd.plot(kind = 'bar',title = '分阵地入网量对比',ax = axes[1,1])
    plt.figure()
    plot_zrw.plot(kind = 'bar',title = '分月份入网量')
    
    plt.show()
    #输出文件
    rst_add_qs=pd.concat([rst_qs,rst],axis = 0)
    rst_add_qs = rst_add_qs.reindex(columns = col)
    rst_add_qs=rst_add_qs.fillna(0)
    rst_add_qs.to_excel(itv_rst_path)

def main():
    read_itv()
    pdqs()
    process()

if __name__=='__main__':
    main()


