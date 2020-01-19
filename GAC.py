import pandas as pd
import os
from openpyxl import load_workbook
Data = input('Data: ')
ObjectName = input('object：')
Radar=input('CR OR FR ')
path = ('F:\\Pythonlearning\\Jenkins Report\\%s\\%s\\%s'%(ObjectName,Data,Radar))

filelist = os.listdir(path)
fillist = os.listdir('F:\\Pythonlearning\\PowerBI_SourceFile\\%s\\%s\\%s'%(ObjectName,Data,Radar))

#先对一个xlsx进行操作，截取一个表中想要的列
for fil in fillist:
    olddi = os.path.join(path, fil)
    if Data in fil:
        os.remove(olddi)
for files in filelist:
#CR的遍历对照结果
    olddir = os.path.join(path,files)
    data = pd.read_excel(olddir,sheetname='PivotReqSrc')

    new_dir = 'F:\\Pythonlearning\\PowerBI_SourceFile\\%s\\%s\\%s\\%s'%(ObjectName,Data,Radar,Data)+files
    writer = pd.ExcelWriter(new_dir)#将该sheet导出为excel文件
    data.to_excel(writer,float_format='%.5f')
    writer.save()
    wb = load_workbook(new_dir)
    ws = wb['Sheet1']
    ws.title = Data
    wb.save(new_dir) #保存变更
    wb.close()
