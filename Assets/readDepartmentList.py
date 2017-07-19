import xlrd
import xlwt
import pandas as pd

if __name__=="__main__":
    excelFilePath=u"2016年12月31日前教科（房娟老师给的表格）.xlsx"
    dataFrame=pd.read_excel(excelFilePath,sheetname="Sheet0",skiprows=[0,1])
    #assetOfCSADataFrame=dataFrame[dataFrame[u"使用单位"]==u"计算机学院/软件职业技术学院"]
    departmentSeries=dataFrame[u"使用单位"].unique()
    departmentList=departmentSeries.tolist()
    #print assetOfCSADataFrame.head()
    #for item in departmentList:
    #    print item,"#",
    #print ""
    print len(departmentSeries)
