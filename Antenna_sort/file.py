# -*- coding: utf-8 -*-
#############################
#
#  Function : Multi-Excel Integrated to Single Excel
#  Author: Johnny Huang
#  Date: 2020/8/4
#
#############################

import pandas as pd
import os

#-----------------------------
#  Function: Read T1:將同一天線不同SKU的結果整理成同一份資料，以方便驗證
#-----------------------------
def antenna(dirPath):
    result = [f for f in os.listdir(dirPath) if os.path.isfile(os.path.join(dirPath, f)) and f!='file.py' ]
    read_id=pd.read_excel(result[0])
    read_id.columns.values[0]=''
    result_T1Port1=read_id.iloc[:,0]
    result_T1Port1=result_T1Port1.rename('')
    result_T1Port2=result_T1Port1
    result_T2Port1=result_T1Port1
    result_T2Port2=result_T1Port1
    result_T3Port1=result_T1Port1
    result_T3Port2=result_T1Port1

    #現有6之不同天線分別是T1 Port 1、T1 Port 2、T2 Port 1、T2 Port 2、T3 Port 1、T3 Port 2
    for i in range(0,len(result)):
        read=pd.read_excel(result[i])
        if (read['T1 Port 1'].isnull().any()==False):
            read_T1Port1=read.loc[:,'T1 Port 1']
            read_T1Port1=read_T1Port1.rename(result[i])
            result_T1Port1=pd.concat([result_T1Port1,read_T1Port1],1)
        if (read['T1 Port 2'].isnull().any()==False):
            read_T1Port2=read.loc[:,'T1 Port 2']
            read_T1Port2=read_T1Port2.rename(result[i])
            result_T1Port2=pd.concat([result_T1Port2,read_T1Port2],1)
        if (read['T2 Port 1'].isnull().any()==False):
            read_T2Port1=read.loc[:,'T2 Port 1']
            read_T2Port1=read_T2Port1.rename(result[i])
            result_T2Port1=pd.concat([result_T2Port1,read_T2Port1],1)
        if (read['T2 Port 2'].isnull().any()==False):
            read_T2Port2=read.loc[:,'T2 Port 2']
            read_T2Port2=read_T2Port2.rename(result[i])
            result_T2Port2=pd.concat([result_T2Port2,read_T2Port2],1)
        if (read['T3 Port 1'].isnull().any()==False):
            read_T3Port1=read.loc[:,'T3 Port 1']
            read_T3Port1=read_T3Port1.rename(result[i])
            result_T3Port1=pd.concat([result_T3Port1,read_T3Port1],1)
        if (read['T3 Port 2'].isnull().any()==False):
            read_T3Port2=read.loc[:,'T3 Port 2']
            read_T3Port2=read_T3Port2.rename(result[i])
            result_T3Port2=pd.concat([result_T3Port2,read_T3Port2],1)

    #該目錄建立data資料夾
    if not os.path.isdir(dirPath+"\data"):
        os.mkdir(dirPath+"\data")
    os.chdir(dirPath+"\data")

    # 輸出各支天線之整理結果
    result_T1Port1.to_excel('T1Port1.xls',index=False)
    result_T1Port2.to_excel('T1Port2.xls',index=False)
    result_T2Port1.to_excel('T2Port1.xls',index=False)
    result_T2Port2.to_excel('T2Port2.xls',index=False)
    result_T3Port1.to_excel('T3Port1.xls',index=False)
    result_T3Port2.to_excel('T3Port2.xls',index=False)

#-----------------------------
#  Main: Read T1:
#-----------------------------
if __name__=='__main__':
    path= os.getcwd() #將excel所在的工作路徑帶入
    antenna(path)