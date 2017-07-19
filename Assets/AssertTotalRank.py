# -*- coding:utf-8 -*-
#计算不同院系的固定资产
import xlrd
import xlwt
import pandas as pd

def getDepartmentList(dataFrame):
    departmentSeries=dataFrame[u"使用单位"].unique()
    departmentList=departmentSeries.tolist()
    return departmentList

def getTotalAssetMoneyOfDepartment(dataFrame,department):
    moneySerise=dataFrame[dataFrame[u"使用单位"]==department][u"价值(元)"]
    totalMoney=moneySerise.sum()
    return totalMoney/10000
    
if __name__=="__main__":
    excelFilePath=u"2016年12月31日前教科（房娟老师给的表格）.xlsx"
    dataFrame=pd.read_excel(excelFilePath,sheetname="Sheet0",skiprows=[0,1])
    departmentList=getDepartmentList(dataFrame)
    #total=getTotalAssetMoneyOfDepartment(dataFrame,u"机电工程学院")
    #print total/10000,u"万"
    departmentMoneyDict=[]
    for oneDepartment in departmentList:
        oneDepartmentMoney={}
        oneDepartmentMoney["department"]=oneDepartment
        oneDepartmentMoney["money"]=getTotalAssetMoneyOfDepartment(dataFrame,oneDepartment)
        departmentMoneyDict.append(oneDepartmentMoney)
    #print departmentMoneyDict
    #新的结果计算出来了
    finalDataFrame=pd.DataFrame(departmentMoneyDict)
    sortedFinalDataFrame=finalDataFrame.sort_values(["money",], ascending=[False,])
    print "Save to excel file !"
    writer=pd.ExcelWriter("output.xlsx")
    sortedFinalDataFrame.to_excel(writer,"Sheet")
    writer.save()
    
