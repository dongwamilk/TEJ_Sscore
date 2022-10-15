# -*- coding: utf-8 -*-
"""
Created on Mon Jul 25 15:12:51 2022

@author: user
"""

import pandas as pd
import numpy as np

Company_basic = pd.read_csv('D:/user/研究所/數值考試資料/CompanyBasic.txt', sep="\t" , encoding='big5hkscs')
df_income = pd.read_csv('D:/user/研究所/蹤影/esg/營收2.csv', sep="#", encoding='big5hkscs')
df_certi = pd.read_excel('D:/user/研究所/蹤影/esg/TESG指標.xlsx', sheet_name = '認證(級別型)')
df_flow = pd.read_excel('D:/user/研究所/蹤影/esg/TESG指標.xlsx', sheet_name = '員工流動率(連續型)')
df_baby = pd.read_excel('D:/user/研究所/蹤影/esg/TESG指標.xlsx', sheet_name = '育嬰留停(連續型)')
df_risk = pd.read_excel('D:/user/研究所/蹤影/esg/TESG指標.xlsx', sheet_name = '失能傷害風險(連續型)')
df_donate = pd.read_excel('D:/user/研究所/蹤影/esg/TESG指標.xlsx', sheet_name = '捐贈(連續型)')

#財報月營收
Income = df_income[df_income['月份'] == 12]
Income_2015 = Income[Income['年月'] > 201412]
Income_2015['年月'] = Income_2015['年月'].astype(str).str[0:4]
Income_2015.rename(columns={'年月': '年'}, inplace=True)


#公司基本資料處理
Company_basic = Company_basic[['公司','簡稱','上市別','TSE 產業別','TSE新產業_名稱','最近上市日']]
Company_basic['公司'] = Company_basic['公司'].str.strip()
Company_basic['簡稱'] = Company_basic['簡稱'].str.strip()
Company_basic['上市別'] = Company_basic['上市別'].str.strip()
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].str.strip()
Company_basic.rename(columns={'公司': '證券代碼'}, inplace=True)
Company_basic = Company_basic[Company_basic['TSE 產業別'] != '  ']

#SASB大轉換
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['紡織纖維','貿易百貨'],'消費品')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['水泥工業','塑膠工業','鋼鐵工業','油電燃氣業','玻璃陶瓷'],'提煉與礦產加工')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['金融業'],'金融')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['食品工業','農林漁牧'],'食品與飲料')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['生技醫療'],'醫療保健')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['建材營造'],'公共建設')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['造紙工業'],'可再生資源與替代能源')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['化學工業','橡膠工業'],'資源轉化')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['文化創意業','資訊服務業','電子商務','觀光事業','社會企業'],'服務')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['電機機械','電器電纜','半導體','電腦及週邊','光電業','電子零組件','電子通路業','其他電子業','農業科技'],'科技與通訊')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['汽車工業','航運業'],'運輸')
Company_basic['TSE新產業_名稱'] = Company_basic['TSE新產業_名稱'].replace(['其他'],'未分類')
Company_basic.rename(columns={'TSE新產業_名稱': 'SASB產業名稱'}, inplace=True)

#公司基本資料與財報合併
df_join = pd.merge(Income_2015,Company_basic, on =['證券代碼','簡稱'], how = 'inner')

##################大合併##########################
#認證
df_certi['證券代碼'] = df_certi['證券代碼'].astype(str)
df_certi['年'] = df_certi['年'].astype(str)

#資訊安全認證
Info_certi = df_certi[df_certi['認證類別'] == '資訊安全管理系統']
Info_certi_n = Info_certi.groupby(['簡稱','年']).agg(Info_certi_count = ('認證類別','count'))
All = pd.merge(df_join,Info_certi_n,on = ['簡稱','年'],how = 'left')

#品質管理認證
qulity_certi = df_certi[df_certi['認證類別'] == '品質管理系統']
qulity_certi_n = qulity_certi.groupby(['簡稱','年']).agg(qulity_certi_count = ('認證類別','count'))
All = pd.merge(All,qulity_certi_n,on = ['簡稱','年'],how = 'left')

#食品安全
food_certi = df_certi[df_certi['認證類別'] == '食品安全管理']
food_certi_n = food_certi.groupby(['簡稱','年']).agg(food_certi_count = ('認證類別','count'))
All = pd.merge(All,food_certi_n,on = ['簡稱','年'],how = 'left')

#職業安全
job_certi = df_certi[df_certi['認證類別'] == '職業安全衛生管理系統']
job_certi_n = job_certi.groupby(['簡稱','年']).agg(job_certi_count = ('認證類別','count'))
All = pd.merge(All,job_certi_n,on = ['簡稱','年'],how = 'left')

#員工流動率
df_flow['員工流動率(%)'] = df_flow['員工流動率(%)'].replace(['.'],np.nan)
flow = df_flow[['簡稱','年','員工流動率(%)']]
flow['年'] = flow['年'].astype(str)
All = pd.merge(All,flow,on = ['簡稱','年'],how = 'left')

#育嬰統計
baby = df_baby.replace('.',np.nan)
baby['年'] = baby['年'].astype(str)
baby['證券代碼'] = baby['證券代碼'].astype(str)
All = pd.merge(All,baby,on = ['證券代碼','簡稱','年'],how = 'left')

#捐贈
df_donate['年月日'] = df_donate['年月日'].astype(str).str[0:4]
df_donate.rename(columns={'年月日': '年'}, inplace=True)
df_donate.drop(columns = ['證券代碼','營收'], inplace = True)
df_donate['現金捐贈金額'] = df_donate['現金捐贈金額'].replace('.',np.nan).astype(float)
donate = df_donate.groupby(['簡稱','年']).agg(Donate = ('現金捐贈金額','sum'))
All = pd.merge(All,donate,on = ['簡稱','年'],how = 'left')
All['Donate'].fillna(0, inplace = True)

#捐贈/營收
All['DI ratio'] = All['Donate'].div(All['營業收入淨額'], axis = 0)
All.fillna({'DI ratio' : 0}, inplace = True)

#############缺失值(公司不同年度中位數)#############

#認證
All['Info_certi_count'].fillna(0, inplace = True)
All['qulity_certi_count'].fillna(0, inplace = True)
All['food_certi_count'].fillna(0, inplace = True)
All['job_certi_count'].fillna(0, inplace = True)

#員工流動率,育嬰留停,復職率,復職留任率

def miss_median(x):
    flow_med = x['員工流動率(%)'].median()
    apply_med = x['育嬰留停申請率%(合計)'].median()
    back_med = x['復職率%(合計)'].median()
    stay_med = x['復職留任率%(合計)'].median()
    a = x.fillna({'員工流動率(%)' : flow_med, 
                  '育嬰留停申請率%(合計)' : apply_med, 
                  '復職率%(合計)' : back_med, 
                  '復職留任率%(合計)' : stay_med})
    return a


All_2 = All.groupby(['簡稱'],group_keys = False).apply(miss_median)

#############缺失值(產業內第三四分位數)#############
def miss_quantile(x):
    flow_quantile = x['員工流動率(%)'].quantile(0.25)
    apply_quantile = x['育嬰留停申請率%(合計)'].quantile(0.25)
    back_quantile = x['復職率%(合計)'].quantile(0.25)
    stay_quantile = x['復職留任率%(合計)'].quantile(0.25)
    a = x.fillna({'員工流動率(%)' : flow_quantile, 
                  '育嬰留停申請率%(合計)' : apply_quantile, 
                  '復職率%(合計)' : back_quantile, 
                  '復職留任率%(合計)' : stay_quantile})
    return a


All_3 = All_2.groupby(['SASB產業名稱','年'],group_keys = False).apply(miss_quantile)
All_3.dropna(axis = 0, how = 'any', inplace = True)

#將越小越好的變數正負號轉換
All_3['員工流動率(%)'] = -All_3['員工流動率(%)']

#############算分#############

#連續型函數
def continuous (x, columns, new_columns):
    list_1 = []
    seri = x[columns]
    for i in range(0,(len(seri))):
        TF_1 = seri.iloc[i] > seri # 比較該值是否大於序列內元素(True,False)
        TF_2 = seri.iloc[i] == seri # 比較該值是否等於序列內元素(True,False)
        n1 = len(list(filter(lambda x: x == True, TF_1))) #該值大於序列內元素之個數(TF_1 true的數量)
        n2 = len(list(filter(lambda x: x == True, TF_2))) - 1 #該值等於序列內元素之個數(TF_2 true的數量-1,因為不含自己)
        A = n1 + n2/2
        B = len(x)
        list_1.append((A/B) * 100)
    x[new_columns] = list_1
    return x

#級別型函數
def step (x, columns, new_columns):
    list_1 = []
    seri_all = x[columns] #產業內所有的公司
    seri = x[x[columns] != 0][columns] #產業內值不為0的公司
    B = len(x[x[columns] != 0]) #產業內值不為0的家數
    if len(seri) > 0: #判斷分子分母不為0
        for i in range(0,(len(seri_all))):
            if seri_all.iloc[i] == 0:
                list_1.append(0)
            else:
                TF_1 = seri_all.iloc[i] > seri # 比較該值是否大於有值序列內元素(True,False)
                TF_2 = seri_all.iloc[i] == seri # 比較該值是否等於有值序列內元素(True,False)
                n1 = len(list(filter(lambda x: x == True, TF_1))) #該值大於大於序列內元素之個數(TF_1 true的數量)
                n2 = len(list(filter(lambda x: x == True, TF_2))) - 1 #該值等於大於序列內元素之個數(TF_2 true的數量-1,因為不含自己)
                A = n1 + n2/2
                list_1.append((A/B) * 100)
        x[new_columns] = list_1
    else :
        x[new_columns] = 0
    return x


#連續型
All_4 = All_3.groupby(['SASB產業名稱','年']).apply(continuous, columns = '員工流動率(%)',new_columns = 'flow').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = '育嬰留停申請率%(合計)',new_columns = 'baby_apply').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = '復職率%(合計)',new_columns = 'baby_back').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = '復職留任率%(合計)',new_columns = 'baby_stay').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = 'DI ratio',new_columns = 'DI_score').reset_index(drop = True)

#級別型
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(step, columns = 'Info_certi_count',new_columns = 'Info_certi').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(step, columns = 'qulity_certi_count',new_columns = 'qulity_certi').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(step, columns = 'food_certi_count',new_columns = 'food_certi').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(step, columns = 'job_certi_count',new_columns = 'job_certi').reset_index(drop = True)

#議題總分(人權、產品品質、員工資訊、員工健康)
All_4['human_right'] = All_4['DI_score']
All_4['product_qulity']  = All_4['qulity_certi'] + All_4['food_certi']
All_4['employees_information'] = All_4['flow']
All_4['employees_health']  = All_4['baby_apply'] + All_4['baby_back'] + All_4['baby_stay'] + All_4['job_certi']

#議題總分再用連續型方法進行排序(議題分數)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = 'human_right',new_columns = 'HUMAN_RIGHT').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = 'product_qulity',new_columns = 'PRODUCT_QULITTY').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = 'employees_information',new_columns = 'EMPLOYEES_INFORMATION').reset_index(drop = True)
All_4 = All_4.groupby(['SASB產業名稱','年']).apply(continuous, columns = 'employees_health',new_columns = 'EMPLOYEES_HEALTH').reset_index(drop = True)

#支柱分數 = 各議題可量化分數平均 * 0.75 + 各議題揭露分數 * 0.25
All_4['S_Score'] = ((All_4['HUMAN_RIGHT'] + All_4['PRODUCT_QULITTY'] + All_4['EMPLOYEES_INFORMATION'] + All_4['EMPLOYEES_HEALTH'])/3) * 0.75


#權重
Ws_dict = {'消費品' : 0.42,'提煉與礦產加工' : 0.25,'金融' : 0.42,
           '食品與飲料' : 0.39,'醫療保健' : 0.53,'公共建設' : 0.3,
           '可再生資源與替代能源' : 0.23,'資源轉化' : 0.3,'服務' : 0.48,
           '科技與通訊' : 0.38,'運輸' : 0.29,'未分類' : 0.33} 

All_4['Ws'] = All_4['SASB產業名稱'].map(Ws_dict)

#Ws_Score
All_4['Ws_Score'] = All_4['S_Score'] * All_4['Ws']



#匯出檔案
All_4.to_excel('Ws_Score_py.xlsx')











