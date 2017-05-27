#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat May 27 18:06:54 2017

@author: jiangjinjin
"""
def processValuationTable(fileDir):
    import re
    import pandas as pd
    import os
    import datetime
    files = os.listdir('{}/'.format(fileDir))
    if not os.path.exists('{}/处理后估值表'.format(fileDir)):
        os.mkdir('{}/处理后估值表'.format(fileDir))
    fileList = []
    for i in files:
        if i[-3:] == 'xls':
            fileList.append(i)
    for fileName in fileList:
        currentDate = datetime.datetime.strptime(re.match('[0-9]+', fileName)[0], '%y%m%d').date()
        data = pd.read_excel("{}/{}".format(fileDir,fileName), sheetname = '今日净值', skiprows = 4)
        data.columns = list(range(data.shape[1]))
        data = data[[1,3]]
        a = []
        b = []
        for i in range(len(data[1])):
            j = data[1][i]
            if type(j) == str and re.match('[0-9]+', j) and len(re.findall('[\u4e00-\u9fa5]+', j)) and data[3][i] == data[3][i]:
                a.append(re.match('[0-9a-zA-Z\u4e00-\u9fa5]+', j)[0])
                b.append(data[3][i])
        outData = pd.DataFrame([[currentDate] * len(a), a, b], index = ['日期', '债券名','数量']).T
        outData.to_excel("{}/处理后估值表/{}.xlsx".format(fileDir,currentDate), index = False)
        
if __name__ == "__main__":
    #父文件夹名
    filename = "估值表"
    processValuationTable(filename)
