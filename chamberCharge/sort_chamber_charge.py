# -*- coding: utf-8 -*-
"""
Created on Tue Oct 20 15:03:03 2020

@author: Johnny_Huang
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

class FileSort(): #整理file，把要檢視的bu從df裡挑出來
    def __init__(self,df,bu):
        self.bg=self.file_sort(df,bu)

    def file_sort(self,df,bu):
        df=df.sort_values(['Hours'],ascending=False)
        fliter=df['★『Please input your BU』'].isin(bu)
        df=df[fliter]
        return df

#sas.iloc[0,2]=sas.iloc[1,7]
#sas=sas.append(pd.Series(), ignore_index=True) 新增空白行eset
#df.to_csv("std.csv",encoding="gbk",index=False ) #写入到csv时，不要将索引写入index = False
class BGclass(FileSort):
    def __init__(self,file_csv,bu):
        super().__init__(file_csv,bu)
        self.aten=self.bg[self.bg['資源'].isin(['天線遠場三維量測實驗室(11F A-Ten Lab) LTE-OTA 時段 / 5G Sub6 時段'])]
        self.aten_5G=self.aten[self.aten['★『Please select LTE OTA or 5G Sub6 or Passive』。'].isin(['5G Sub6'])]
        self.aten_4G=self.aten[self.aten['★『Please select LTE OTA or 5G Sub6 or Passive』。'].isin(['LTE OTA','Passive'])]
        self.catr=self.bg[self.bg['資源'].isin(['縮距場毫米波天線量測系統 (4F CATR lab)'])]
        self.sort_df=self._concat_file(self._concat_file(self.aten_5G,self.aten_4G),self.catr)
        self.aten_sum_5G=self.hourSum(self.aten_5G)
        self.aten_sum_4G=self.hourSum(self.aten_4G)
        self.catr_sum=self.hourSum(self.catr)
        self.BU=self.buSum(self.bg,bu)
        self.add_hour_to_file(bu,self.BU)
        self.aten=self._concat_file(self.aten_5G,self.aten_4G)
        self.BG=self._concat_file(self.aten,self.catr)

    # def rowNum(self,chamber): #計算列數
    #     return chamber.shape[0]-1

    def hourSum(self,chamber): #計算用了幾小時
        return chamber['Hours'].sum()

    def add_hour_to_file(self,bu_list,bu_listSum): #將分類後的df加入各catr與aten的使用時數
        self.catr=self.catr.append(pd.Series([np.nan]),ignore_index=True)
        self.aten_5G=self.aten_5G.append(pd.Series([np.nan]),ignore_index=True)
        self.aten_4G=self.aten_4G.append(pd.Series([np.nan]),ignore_index=True)
        self.catr.iloc[self.catr.shape[0]-1,3]= 'catr 總使用時數:'
        self.catr.iloc[self.catr.shape[0]-1,4]= self.catr_sum
        self.aten_5G.iloc[self.aten_5G.shape[0]-1,3]= 'Aten 5G總使用時數:'
        self.aten_5G.iloc[self.aten_5G.shape[0]-1,4]= self.aten_sum_5G
        self.aten_4G.iloc[self.aten_4G.shape[0]-1,3]= 'Aten 4G總使用時數:'
        self.aten_4G.iloc[self.aten_4G.shape[0]-1,4]= self.aten_sum_4G

        #將df加入各bu的使用時數
        index_num=0
        for i in range(0,len(bu_list)):
            self.aten_4G.iloc[self.aten_4G.shape[0]-1,5+i+index_num]=bu_list[i]
            self.aten_4G.iloc[self.aten_4G.shape[0]-1,6+i+index_num]=bu_listSum['4G'][bu_list[i]]
            self.aten_5G.iloc[self.aten_5G.shape[0]-1,5+i+index_num]=bu_list[i]
            self.aten_5G.iloc[self.aten_5G.shape[0]-1,6+i+index_num]=bu_listSum['5G'][bu_list[i]]
            self.catr.iloc[self.catr.shape[0]-1,5+i+index_num]=bu_list[i]
            self.catr.iloc[self.catr.shape[0]-1,6+i+index_num]=bu_listSum['catr'][bu_list[i]]
            index_num+=1

    def buSum(self,bg,bu):#計算各bg中的bu分別使用了多少時數
        aten4G_list={}
        aten5G_list={}
        catr_list={}
        for k in range(0,len(bu)):
            sum_4G=0; sum_5G=0; catr=0
            for i in range(0,bg.shape[0]):
                if (bg.iloc[i,21]==bu[k] and bg.iloc[i,0]=="天線遠場三維量測實驗室(11F A-Ten Lab) LTE-OTA 時段 / 5G Sub6 時段"
                    and (bg.iloc[i,22]=='LTE OTA' or bg.iloc[i,22]=='Passive')):
                    sum_4G+=bg.iloc[i,4]
                elif (bg.iloc[i,21]==bu[k] and bg.iloc[i,0]=="天線遠場三維量測實驗室(11F A-Ten Lab) LTE-OTA 時段 / 5G Sub6 時段"
                    and (bg.iloc[i,22]=='5G Sub6')):
                    sum_5G+=bg.iloc[i,4]
                elif (bg.iloc[i,21]==bu[k] and bg.iloc[i,0]=="縮距場毫米波天線量測系統 (4F CATR lab)"):
                    catr+=bg.iloc[i,4]

            aten4G_list[bu[k]]=sum_4G
            aten5G_list[bu[k]]=sum_5G
            catr_list[bu[k]]=catr
        return {'4G':aten4G_list,'5G':aten5G_list,'catr':catr_list}

    def _concat_file(self,file1,file2): #合併df
        return pd.concat([file1,file2],axis=0,ignore_index = True )


if __name__=='__main__':
    df=pd.read_csv('chamber.csv',encoding='Big5')
    # df 裡面多了"參與者"
    # print(len(df['開始']))
    # print(df['開始'])
    for i in range(49):
        if df['開始'][i][0] == '2' and df['開始'][i][1] == '0' and df['開始'][i][2] == '0' :
            df['開始'][i] =df['開始'][i][3:]
        else:
            print("not found")
        if df['結束'][i][0] == '2' and df['結束'][i][1] == '0' and df['結束'][i][2] == '0' :
            df['結束'][i] = df['結束'][i][3:]
        else:
            print("not found")

    BG_sas=BGclass(df,['SAS_JER500'])
    sas=BG_sas.BG
    BG_ais=BGclass(df,["AIC_T50000","ICS_T30000","SMA_T20000"])
    ais=BG_ais.BG
    BG_ch=BGclass(df,["CH1_H60000","CH2_H50000","CH3_H70000"])
    ch=BG_ch.BG
    BG_nw=BGclass(df,["NW1_D10000","NW2_D20000","NW3_D30000"])
    nw=BG_nw.BG
    BG_atd=BGclass(df,['JN0000'])
    atd=BG_atd.BG

    writer = pd.ExcelWriter('allBU.xlsx') # pylint: disable=abstract-class-instantiated

    sas.to_excel(writer,index=False,sheet_name='SAS')
    ais.to_excel(writer,index=False,sheet_name='AIS')
    ch.to_excel(writer,index=False,sheet_name='Connected Home')
    nw.to_excel(writer,index=False,sheet_name='Net Working')
    atd.to_excel(writer,index=False,sheet_name='ATD')
    writer.save()

    book = load_workbook(r'BUcost.xlsx')
    writer = pd.ExcelWriter('BUcost.xlsx', engine='openpyxl') # pylint: disable=abstract-class-instantiated
    writer.book = book
    sas.to_excel(writer,index=False,sheet_name='SAS')
    ais.to_excel(writer,index=False,sheet_name='AIS')
    ch.to_excel(writer,index=False,sheet_name='Connected Home')
    nw.to_excel(writer,index=False,sheet_name='Net Working')
    atd.to_excel(writer,index=False,sheet_name='ATD')
    writer.save()

#############################
# Example how to use:
# df=pd.read_csv('chamber.csv',encoding='Big5')
# BG_nw=BGclass(df,["NW1_D10000","NW2_D20000","NW3_D30000"])
# nw=BG_nw.BG
# print(nw)
############################


#https://blog.csdn.net/hnxyyzx/article/details/106245046