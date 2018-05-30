# -*- coding: utf-8 -*-
"""
Created on Thu Mar 29 11:00:56 2018

@author: bin
"""
import numpy as np
import pandas as pd
import time
import os

global osc_new, drdd, hhr, kg, znzw, schk, zwyg, \
    yhq, kg_jgl, sjzs,sssl,osc_tt

#####################准备阶段######################
#受理订单文件名
last_day = time.localtime(time.time()-24*3600)
time_name = time.strftime('%Y-%m-%d',last_day)
osc_fname = ''.join(['受理订单_',time_name,'.xlsx'])
#路径
main_file = '002-new_data'
osc_new_path = os.path.join(main_file,osc_fname)
osc_old_path = os.path.join(main_file,'001-OSC_shouli.csv')
drdd_path = os.path.join(main_file,'002-导入清单统计.xls')
hhr_path=os.path.join(main_file,'003-dwddqd.csv')
others_path =  os.path.join(main_file,'004-Others.xlsx')
#mid&old_data
mid_file = '001-mid&old_data'
DRD_path = os.path.join(mid_file,'001-DailyReport_data.csv')
DRH_path = os.path.join(mid_file,'002-DR_history.xlsx')
#result
rst_file = '003-result'
combined_path = os.path.join(rst_file, '007-combined.xlsx')
#本月
the_month = time.strftime('%Y-%m',last_day)
#######################数据读入######################
def read_xt():
    global osc_new, drdd, hhr, kg, znzw, schk, zwyg, \
    yhq, kg_jgl, sjzs,sssl,osc_tt
    #OSC系统订单
    try:
        osc_new = pd.read_excel(osc_new_path)
        osc_new['订单创建时间']=pd.to_datetime(osc_new['订单创建时间'])
        osc_new.set_index('订单创建时间', inplace=True)
        osc_new=osc_new[the_month].reset_index()
        osc_old = pd.read_csv(osc_old_path,encoding='GBK')
        osc_tt = pd.concat([osc_new,osc_old])
        osc_col = osc_new.columns
        osc_tt = osc_tt.reindex(columns=osc_col)
        drdd = pd.read_excel(drdd_path)
    except:
        print('OSC数据文件不存在或格式不符，请核查清楚')
    #其他数据
    try:
        hhr = pd.read_csv(hhr_path,encoding='GBK')#合伙人
        kg = pd.read_excel(others_path,sheetname = '快供')
        kg_jgl = pd.read_excel(others_path,sheetname = '快供竣工率')
        znzw = pd.read_excel(others_path,sheetname = '智能组网',dtype ={'分公司':str})
        sjzs = pd.read_excel(others_path,sheetname = '升级助手')
        schk = pd.read_excel(others_path,sheetname = '商城换卡')
        zwyg = pd.read_excel(others_path,sheetname = '装维翼购')
        yhq = pd.read_excel(others_path,sheetname = '优惠券')
        sssl = pd.read_excel(others_path,sheetname = '提速实时受理')
    except:
        print('非OSC数据文件不存在或格式不符，请核查清楚')
#####################判断数据日期###################
def pdrq():
    global osc_new, drdd, hhr, kg, znzw, schk, zwyg, yhq
    txt1 = 'OSC数据'+str(osc_new['订单创建时间'].max())[:10]
    txt2 = 'OSC导入统计'+str(drdd['下单时间'].max()[:10])
    txt3 = '合伙人'+str(hhr['下单时间'].max()[:10])
    txt4 = '快供'+str(kg['下单日期'].max())
    txt5 = '智能组网'+str(znzw['下单时间'].max()[:10])
    txt6 = '商城换卡'+str(schk['下单时间'].max())
    txt7 = '装维翼购'+str(zwyg['下单日期'].max())
    txt8 = '优惠券'+str(yhq['下单日期'].max())
    txt = '\n'.join([txt1,txt2,txt3,txt4,txt5,txt6,txt7,txt8])
    return txt
#####################数据处理######################

def process():
    global osc_new, drdd, hhr, kg, znzw, schk, zwyg, yhq, sjzs, sssl
    key_lst = ['分公司','下单日期','是否受理成功', '来源小类', '场景名称','场景细类','业务小类']
    
    ###osc_tt数据处理
    osc_tt.drop_duplicates('订单编号',inplace = True)
    osc_tt.to_csv(osc_old_path,index=None)#把当月累计数据保存为'001-OSC_shouli.csv'
    osc_tt.rename(columns = {'订单状态':'受理状态','CRM状态':'订单状态','订单城市':'分公司'},inplace = True)
    osc_tt['下单日期']=pd.to_datetime(osc_tt['订单创建时间']).dt.strftime('%Y%m%d')
    #osc_tt受理状态
    osc_tt['是否受理成功']='否'
    osc_tt.loc[osc_tt['受理状态'].isin(['待受理','待审核']),'是否受理成功']='超时'
    osc_tt.loc[osc_tt['受理状态'].isin(["订单已完结","待受理归档"]),'是否受理成功']='是'
    #osc_tt竣工状态
    osc_tt['是否竣工']=0
    osc_tt.loc[(osc_tt['订单状态'].notnull())&(osc_tt['订单状态'].str.contains('竣工')),'是否竣工']=1
    osc_tt.loc[(osc_tt['外呼状态'].notnull())&(osc_tt['外呼状态'].isin(['外呼成功','无需外呼'])),'是否竣工']=1
    #osc_tt来源小类
    osc_tt['来源小类']=osc_tt['订单来源']
    ######################OSC系统新增活动#########################
    #osc_tt场景
    osc_tt['场景名称']=np.nan
    osc_tt['业务小类']=np.nan
    osc_tt.loc[osc_tt['推广渠道']=='B2I-订单流转','场景名称']='集团订单平台'
    osc_tt.loc[osc_tt['推广渠道']=='1805ZJ159SJBXL','场景名称']='159升级不限量'
    osc_tt.loc[osc_tt['推广渠道']=='政企OMO活动','场景名称']='进政企做个人'
    osc_tt.loc[osc_tt['订单来源']=='O2O_10000号','场景名称']='O2O_10000号'
    osc_tt.loc[osc_tt['工作组名称']=='汕尾10000客服中心-外呼资源核查工作组','场景名称']='外呼稽核'
    osc_tt.loc[(osc_tt['工作组名称'].notnull())&osc_tt['工作组名称'].str.contains('单进融'),'场景名称']='单进融'
    osc_tt.loc[osc_tt['工作组名称']=='省公司电子渠道运营中心-京东淘宝','场景名称']='517订单'
    #osc_tt业务小类
    osc_tt.loc[osc_tt['场景名称']=='集团订单平台','业务小类']='集团B2I'
    osc_tt.loc[osc_tt['场景名称']=='159升级不限量','业务小类']='不限量卡'
    osc_tt.loc[osc_tt['场景名称']=='进政企做个人','业务小类']='不限量卡'
    osc_tt.loc[osc_tt['场景名称']=='O2O_10000号','业务小类']='10000号'
    osc_tt.loc[osc_tt['场景名称']=='外呼稽核','业务小类']='10000号'
    osc_tt.loc[osc_tt['场景名称']=='单进融','业务小类']='单进融'
    osc_tt.loc[osc_tt['场景名称']=='517订单','业务小类']='融合'
    osc_tt['场景细类']=osc_tt['场景名称']
    #######################OSC系统新增活动########################
    
    #osc_tt订单汇总处理
    osc_grouped = osc_tt.groupby(key_lst)
    cleaned_osc=osc_grouped.agg({'订单编号':'count','是否竣工':'sum'})
    cleaned_osc.reset_index(inplace=True)
    cleaned_osc.rename(columns={'订单编号':'订单量','是否竣工':'竣工量'},inplace = True)
    
    ###OSC导入订单数据处理
    drdd.rename(columns={'地区':'分公司','系统来源':'来源小类'},inplace=True)#来源小类
    drdd['下单日期']=pd.to_datetime(drdd['下单时间']).dt.strftime('%Y%m%d')
    drdd['是否受理成功']='否'
    drdd.loc[drdd['订单状态'].isin(["订单已完结","待受理归档"]),'是否受理成功']='是'
    drdd['是否竣工']=0
    drdd.loc[drdd['crm状态']=='竣工','是否竣工']=1
    drdd['场景名称']=np.nan
    drdd['业务小类']=np.nan
    drdd.loc[drdd['销售品名称']=='选号','场景名称']='线上选号'
    drdd.loc[(drdd['统计类型']=='3升4')&(drdd['分公司']=='佛山'),'场景名称']='佛山3升4'
    drdd.loc[drdd['场景名称']=='线上选号','业务小类']='选号吧'
    drdd.loc[drdd['场景名称']=='佛山3升4','业务小类']='3升4'
    drdd['场景细类']=drdd['场景名称']
    #OSC导入订单汇总处理
    drdd_grouped = drdd.groupby(key_lst)
    cleaned_drdd=drdd_grouped.agg({'订单号':'count','是否竣工':'sum'})
    cleaned_drdd.reset_index(inplace=True)
    cleaned_drdd.rename(columns={'订单号':'订单量','是否竣工':'竣工量'},inplace = True)
    
    ######其他日报清单处理
    #快供
    kg['下单日期']=kg['下单日期'].astype(str)
    kg['场景名称']='快供平台'
    kg['场景细类']='快供平台'
    kg['是否受理成功']='未知'
    cleaned_kg = pd.DataFrame(kg.groupby(key_lst).sum()['当日订单数'])
    cleaned_kg.reset_index(inplace=True)
    cleaned_kg['下单月份']=cleaned_kg['下单日期'].str.slice(0,6)
    cleaned_kg.rename(columns={'当日订单数':'订单量'},inplace=True)
    kg_jgl['下单月份']=kg_jgl['下单月份'].astype(str)
    cleaned_kg = pd.merge(cleaned_kg,kg_jgl,how='left',on=['下单月份','分公司','场景名称'])
    cleaned_kg['竣工量']=cleaned_kg['订单量']*cleaned_kg['激活率']
    data_lst = ['分公司', '下单日期', '是否受理成功', '来源小类', '场景名称', '场景细类', '业务小类', '订单量', '竣工量']
    cleaned_kg=cleaned_kg[data_lst]
    
    #装维翼购数据
    zwyg = zwyg.groupby(['下单日期','分公司']).sum().stack()
    zwyg=zwyg.reset_index()
    zwyg_cj=zwyg['level_2'].str.split('-',n=2,expand=True)
    zwyg=pd.concat([zwyg,zwyg_cj],axis=1)
    zwyg.columns=['下单日期','分公司','场景细类','订单数','订单类型','业务小类']
    zwyg['订单量']=0
    zwyg['竣工量']=0
    zwyg.loc[zwyg['订单类型']=='当日下单量','订单量']=zwyg['订单数']
    zwyg.loc[zwyg['订单类型']=='当日竣工量','竣工量']=zwyg['订单数']
    zwyg['是否受理成功']='未知'
    zwyg['来源小类']='装维翼购平台'
    zwyg['场景名称']='装维翼购'
    zwyg['场景细类']=zwyg['场景名称'].str.cat(zwyg['业务小类'],sep='-')
    zwyg['下单日期']=zwyg['下单日期'].astype('str')
    cleaned_zw = zwyg.groupby(key_lst).sum()[['订单量','竣工量']]
    cleaned_zw.reset_index(inplace=True)
    cleaned_zw = cleaned_zw.loc[cleaned_zw['分公司']!='全省']
    
    #合伙人数据
    hhr.rename(columns={'市分公司':'分公司','订单类型':'业务小类'},inplace = True)
    hhr=hhr.loc[hhr['分公司']!='-']
    hhr=hhr.loc[~hhr['合伙人'].str.contains('测')]
    hhr=hhr.loc[~hhr['销售品名称'].str.contains('测')]
    hhr['下单日期']=pd.to_datetime(hhr['下单时间']).dt.strftime('%Y%m%d')
    hhr['是否受理成功']='未知'
    hhr['场景名称']='合伙人'
    hhr['来源小类']='合伙人平台'
    hhr['是否竣工']=0
    hhr.loc[hhr['销售品名称'].str.contains('日租卡'),'业务小类']='天翼小白'
    
    hhr_dy={'天翼合伙人移动':'天翼小白',
            '移动号卡':'天翼小白',
            '固话业务':'天翼小白',
            '电视':'天翼高清',
            '宽带':'宽带新装',
            '融合新装':'融合业务',
            '融合加装':'融合业务',
            }
    hhr['业务小类']=hhr['业务小类'].map(lambda x:hhr_dy.get(x,x))
    hhr.loc[hhr['订单状态']=='已竣工','是否竣工']=1
    hhr['场景细类']=hhr['场景名称'].str.cat(hhr['业务小类'],sep='-')
    #合伙人汇总处理
    cleaned_hhr=hhr.groupby(key_lst).agg({'订单编码':'count','是否竣工':'sum'})
    cleaned_hhr.reset_index(inplace = True)
    cleaned_hhr.rename(columns={'订单编码':'订单量','是否竣工':'竣工量'},inplace=True)
    
    #优惠券数据
    yhq['下单日期']=yhq['下单日期'].astype(str)
    yhq['是否受理成功']='未知'
    yhq['场景名称']='优惠券统计'
    yhq['场景细类']='优惠券统计'
    yhq['来源小类']='优惠券平台'
    yhq['业务小类']='优惠券'
    cleaned_yhq = yhq
    
    #智能组网
    znzw['下单日期']=pd.to_datetime(znzw['下单时间']).dt.strftime('%Y%m%d')
    znzw['是否受理成功']='否'
    znzw.loc[znzw['订单状态']=='已受理','是否受理成功']='是'
    znzw['场景名称']='智能组网'
    znzw['来源小类']='互联网化订单平台'
    znzw['业务小类']='智能组网'
    znzw['场景细类']='智能组网'
    znzw['是否竣工']=0
    znzw.loc[znzw['IB订单状态'].isin(['完工','已竣工','已完成']),'是否竣工']=1
    cleaned_zn = znzw.groupby(key_lst).agg({'订单号':'count','是否竣工':'sum'})
    cleaned_zn.rename(columns={'订单号':'订单量', '是否竣工':'竣工量'},inplace=True)
    cleaned_zn.reset_index(inplace=True)
    
    #升级助手
    sjzs=sjzs.loc[sjzs['地址']=='当面换卡']
    sjzs['下单日期']=pd.to_datetime(sjzs['注册时间']).dt.strftime('%Y%m%d')
    sjzs['场景名称']='当面换卡'
    sjzs['业务小类']='3升4'
    sjzs['来源小类']='升级助手平台'
    sjzs['场景细类']='当面换卡'
    sjzs['是否受理成功']='未知'
    sjzs['竣工量']=0
    cleaned_sj = sjzs.groupby(key_lst).agg({'ID':'count','竣工量':'sum'})
    cleaned_sj.rename(columns={'ID':'订单量'},inplace=True)
    cleaned_sj.reset_index(inplace=True)
    
    #商城换卡
    schk=schk.loc[(schk['支付方式']== "货到付款")&
                  (schk['状态']!='审核不通过')&
                  (schk['状态']!='已取消')]
    schk['分公司']=schk['订单编码'].str.slice(1,3)
    schk['下单日期']=schk['下单时间'].astype(str)
    schk['场景名称']='商城换卡'
    schk['业务小类']='3升4'
    schk['来源小类']='商城换卡平台'
    schk['场景细类']='商城换卡'
    schk['是否受理成功']='未知'
    cleaned_sc = schk.groupby(key_lst).agg({'订单编码':'count','是否电渠激活':'sum'})
    cleaned_sc.rename(columns={'订单编码':'订单量','是否电渠激活':'竣工量'},inplace=True)
    cleaned_sc.reset_index(inplace=True)
    
    #实时受理
    sssl=sssl.loc[sssl['是否下单']=='已处理']
    sssl['下单日期']=sssl['下单日期'].astype(str)
    sssl['是否受理成功']='是'
    sssl['场景名称']='提速实时受理'
    sssl['业务小类']='提速'
    sssl['来源小类']='实时受理平台'
    sssl['场景细类']='提速实时受理'
    sssl['竣工量']=sssl['订单量']
    cleaned_sl = sssl.groupby(key_lst).agg({'订单量':'sum','竣工量':'sum'})
    cleaned_sl.reset_index(inplace=True)
    
    #####数据汇总处理
    data_lst = [cleaned_osc,cleaned_drdd,cleaned_kg,cleaned_zw,cleaned_hhr,cleaned_yhq,cleaned_zn,cleaned_sj,cleaned_sc,cleaned_sl]
    cleaned_data = pd.concat(data_lst)
    cleaned_data['下单月份']=cleaned_data['下单日期'].str.slice(0,6)
    cleaned_data['下单日']=cleaned_data['下单日期'].str.slice(-2,)
    #####################数据输出######################
    output_lst=['分公司',
     '下单日期',
     '是否受理成功',
     '场景名称',
     '业务小类',
     '来源小类',
     '提速细类',
     '场景细类',
     '订单量',
     '竣工量',
    '下单月份',
    '下单日']
    
    rst = cleaned_data.reindex(columns=output_lst)
    rst.to_csv(DRD_path,index=False)
    
    #合并历史数据
    htry=pd.read_excel(DRH_path,dtype={'分公司':str,'下单日':str})
    combined = pd.concat([htry,rst])
    combined.reindex(columns = output_lst)
    combined.to_excel(combined_path,index = False)
def main():
    read_xt()
    pdrq()
    process()
    
if __name__=='__main__':
    main()
#==============================================================================
# © 2018 Zheng, All Rights Reserved
#==============================================================================
