# -*- coding:utf-8 -*-
#计算不同院系的固定资产
import xlrd
import xlwt
import pandas as pd

#获取属于某一个单位的固定资产，返回值是DataFrame
def getAssetDataFrameOfDepartment(dataFrame,department):
    assetOfDepartment=dataFrame[dataFrame[u"使用单位"]==department]
    return assetOfDepartment

#以列表的形式返回获取所有的"购置日期"，返回前去除重复
def getYearListDataFrameOfAsset(dataFrame):
    yearList=dataFrame[u"购置日期"].sort_values().unique().tolist()
    return yearList

#指定某一个年份，格式为u"2014",并返回在固定资产"购置之日"属性中年份中所有大于该年份的日期列表
def yearListAfterSpecialYear(yearList,specialYear):  
    theYear=int(specialYear)
    finalYearList=[]
    for oneYear in yearList:
        oneYearInArrayFormat=oneYear.split(u"-")
        yearParts=int(oneYearInArrayFormat[0])
        if yearParts>=theYear:
            finalYearList.append(oneYear)
    return finalYearList
    
#指定一个"采购日期"列表，返回属于该列表的固定资产价格总数（万元），同时打印每个采购日期对应的固定资产数额
def totalMoneyAtYear(dataFrame,yearList):
    totalMoney=0  #以万为单位计算
    for oneYear in yearList:
        selectDataFrame=dataFrame[dataFrame[u"购置日期"]==oneYear]
        moneySerise=selectDataFrame[u"价值(元)"]
        totalMoney=totalMoney+moneySerise.sum()/10000
        print u"时间:{0}\t 本阶段资产增加:{1}万元".format(oneYear,moneySerise.sum()/10000)
    return totalMoney

#筛选属于计算机学院的资产数据，并保存为output.xlsx文件
if __name__=="__main__":
    excelFilePath=u"2016年12月31日前教科（房娟老师给的表格）.xlsx"
    dataFrame=pd.read_excel(excelFilePath,sheetname="Sheet0",skiprows=[0,1])
    assetOfCSADataFrame=getAssetDataFrameOfDepartment(dataFrame,u"计算机学院/软件职业技术学院")
    yearList= getYearListDataFrameOfAsset(assetOfCSADataFrame)
    
    
    selectPurchaseYearList=yearListAfterSpecialYear(yearList,"2014") #制定年份，返回在此年份后的购置日期列表
    totalMoney=totalMoneyAtYear(assetOfCSADataFrame,selectPurchaseYearList)
    print u"上述时间段内，固定资产增加总数为{}万元".format(totalMoney)
    
    print u"生产在该时间段内的固定资产清单"
    #print selectPurchaseYearList
    assetOfCSAPurchaseAtDateDataFrame=assetOfCSADataFrame[assetOfCSADataFrame[u"购置日期"].isin(selectPurchaseYearList)]
    
    #将处理后的数据放到新的表格中
    newIndexList=range(1,(assetOfCSAPurchaseAtDateDataFrame[u"序号"].count()+1)) #重新生产物品排序索引
    newIndexSerise=pd.Series(newIndexList,name="序号")
    assetOfCSAPurchaseAtDateOfResetIndexDataFrame=assetOfCSAPurchaseAtDateDataFrame.reset_index()
    assetOfCSAPurchaseAtDateOfResetIndexDataFrame[u"序号"]=newIndexSerise
    
    print u"正在保存!"
    writer=pd.ExcelWriter("output.xlsx")
    assetOfCSAPurchaseAtDateOfResetIndexDataFrame.to_excel(writer,"Sheet",columns=[u"序号",u"使用单位",u"使用部门",u"资产编码",u"资产编码",u"资产分类",u"资产名称",u"型号",u"规格",u"价值(元)",u"使用部门",u"资产编码",u"资产编码",u"购置日期",])
    writer.save()
    
