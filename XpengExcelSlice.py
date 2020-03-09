import glob
import pandas as pd
import os
#from openpyxl import load_workbook
time = input('Data: ')
#pro = input('object：')
#radar=input('CR OR FR')
path = ('N:\\Tech\\DA\\03_Radar\\09_Engineering\\58_RM\\GEN5\\RM Status\\Jenkins Report\\Xpeng_ED\\%s\\*'%(time))

#Assign the stroge path of front radar 
outpath_FRCust='N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/CustRS/'
outpath_FRSW_1R1V  = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/SWRS/1R1V_3R1V/'
outpath_FRSW_5R1V  = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/SWRS/5R1V/'
outpath_FRSys_1R1V = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/SysRS/1R1V_3R1V/'
outpath_FRSys_5R1V = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/SysRS/5R1V/'
outpath_FRSys_Test_1R1V = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/SysTestRS/1R1V_3R1V/'
outpath_FRSys_Test_5R1V = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_FR5/SysTestRS/5R1V/'

#Assign the stroge path of corner radar
outpath_CRCust= 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/CustRS/'
outpath_CRSW_1R1V_RC  = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SWRS/1R1V_3R1V/RC/'
outpath_CRSW_5R1V1D_FC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SWRS/5R1V1D/FC/'
outpath_CRSW_5R1V1D_RC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SWRS/5R1V1D/RC/'
outpath_CRSys_1R1V_RC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SysRS/1R1V_3R1V/RC/'
outpath_CRSys_5R1V1D_FC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SysRS/5R1V1D/FC/'
outpath_CRSys_5R1V1D_RC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SysRS/5R1V1D/RC/'
outpath_CRSys_Test_1R1V_RC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SysTestRS/1R1V_3R1V/RC/'
outpath_CRSys_Test_5R1V1D_RC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SysTestRS/5R1V1D/RC/'
outpath_CRSys_Test_5R1V1D_FC = 'N:/Tech/DA/03_Radar/09_Engineering/58_RM/GEN5/RM Status/Xpeng_ED_CR5/SysTestRS/5R1V1D/FC/'

#get all Excel-files under the folder "path", and make a list for their name
file = glob.glob(path)

#collect the all useful attributes from Power BI into the list 
list_1=['UniqueId','ID','DA_AUTO_CFR_Index',
     'DA_Review_Finding_Xpeng','DA_Review_Status_Xpeng',
      'DA_Safety_Integrity','DA_Security_Relevance',
      'DA_Variant_Implemented','DA_Verification_Criteria','SW_Release_Number_ED',
      'Status ED','LinkedReqTestId ED','RqmTestId ED','TestResult ED']

#make a loop so that all Excel-files can be matched.
for i in file:

      # matchming the assigend Excle-file, keywords combination
    if 'Corner_CustomerReq' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')   #reade the date of sheet_name='PivotReqSrc'
        data = data.loc[:, list_1]  #screen out each useful row, via matchming "list_1"
        New_Name = 'Cust-%s'%(time) # add name by the created time 
        print('importing 1/17.file\n') #print out import-process is going on 

        #deleat the old existed Excel-file
        if os.path.exists(outpath_CRCust + '%s.xlsx'%New_Name):
           os.remove(outpath_CRCust + '%s.xlsx'%New_Name)

        # save the new Excel-file into the assigend path     
        data.to_excel(outpath_CRCust + '%s.xlsx'%New_Name, index=False)

    elif 'FrontCorner_5R1V1D_SoftwareTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'SW-%s'%(time)
        print('importing 2/17.file\n')
        if os.path.exists(outpath_CRSW_5R1V1D_FC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSW_5R1V1D_FC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSW_5R1V1D_FC + '%s.xlsx'%New_Name, index=False)

    elif 'FrontCorner_5R1V1D_SystemReq' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 3/17.file\n')
        if os.path.exists(outpath_CRSys_5R1V1D_FC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSys_5R1V1D_FC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSys_5R1V1D_FC + '%s.xlsx'%New_Name, index=False)

    elif 'FrontCorner_5R1V1D_SystemTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 4/17.file\n')
        if os.path.exists(outpath_CRSys_Test_5R1V1D_FC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSys_Test_5R1V1D_FC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSys_Test_5R1V1D_FC + '%s.xlsx'%New_Name, index=False)

    elif 'RearCorner_1R1V_3R1V_SoftwareTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'SW-%s'%(time)
        print('importing 5/17.file\n')
        if os.path.exists(outpath_CRSW_1R1V_RC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSW_1R1V_RC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSW_1R1V_RC+'%s.xlsx'%New_Name, index=False)

    elif 'RearCorner_1R1V_3R1V_SystemReq' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 6/17.file\n')
        if os.path.exists(outpath_CRSys_1R1V_RC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSys_1R1V_RC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSys_1R1V_RC+'%s.xlsx'%New_Name, index=False)

    elif 'RearCorner_1R1V_3R1V_SystemTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 7/17.file\n')
        if os.path.exists(outpath_CRSys_Test_1R1V_RC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSys_Test_1R1V_RC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSys_Test_1R1V_RC+'%s.xlsx'%New_Name, index=False)

    elif 'RearCorner_5R1V1D_SoftwareTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'SW-%s'%(time)
        print('importing 8/17.file\n')
        if os.path.exists(outpath_CRSW_5R1V1D_RC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSW_5R1V1D_RC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSW_5R1V1D_RC+'%s.xlsx'%New_Name, index=False)

    elif 'RearCorner_5R1V1D_SystemReq' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 9/17.file\n')
        if os.path.exists(outpath_CRSys_5R1V1D_RC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSys_5R1V1D_RC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSys_5R1V1D_RC+'%s.xlsx'%New_Name, index=False)

    elif 'RearCorner_5R1V1D_SystemTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 10/17.file\n')
        if os.path.exists(outpath_CRSys_Test_5R1V1D_RC + '%s.xlsx'%New_Name):
           os.remove(outpath_CRSys_Test_5R1V1D_RC + '%s.xlsx'%New_Name)
        data.to_excel(outpath_CRSys_Test_5R1V1D_RC+'%s.xlsx'%New_Name, index=False)

    #for the front radar 
    elif 'FR_1R1V_3R1V_SoftwareTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'SW-%s'%(time)
        print('importing 11/17.file\n')
        if os.path.exists(outpath_FRSW_1R1V + '%s.xlsx'%New_Name):
           os.remove(outpath_FRSW_1R1V + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRSW_1R1V+'%s.xlsx'%New_Name, index=False)

    elif 'FR_1R1V_3R1V_SystemTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 12/17.file\n')
        if os.path.exists(outpath_FRSys_Test_1R1V + '%s.xlsx'%New_Name):
           os.remove(outpath_FRSys_Test_1R1V + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRSys_Test_1R1V+'%s.xlsx'%New_Name, index=False)

    elif 'FR_5R1V_SoftwareTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'SW-%s'%(time)
        print('importing 13/17.file\n')
        if os.path.exists(outpath_FRSW_5R1V + '%s.xlsx'%New_Name):
           os.remove(outpath_FRSW_5R1V + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRSW_5R1V+'%s.xlsx'%New_Name, index=False)


    elif 'FR_5R1V_SystemTest' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 14/17.file\n')
        if os.path.exists(outpath_FRSys_Test_5R1V + '%s.xlsx'%New_Name):
           os.remove(outpath_FRSys_Test_5R1V + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRSys_Test_5R1V+'%s.xlsx'%New_Name, index=False)


    elif 'Front_1R1V_3R1V_SystemReq' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 15/17.file\n')
        if os.path.exists(outpath_FRSys_1R1V + '%s.xlsx'%New_Name):
           os.remove(outpath_FRSys_1R1V + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRSys_1R1V+'%s.xlsx'%New_Name, index=False)

    elif 'Front_CustomerReqStatus' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Cust-%s'%(time)
        print('importing 16/17.file\n')
        if os.path.exists(outpath_FRCust + '%s.xlsx'%New_Name):
           os.remove(outpath_FRCust + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRCust+'%s.xlsx'%New_Name, index=False)

    elif'Front_5R1V_SystemReqStatus' in i:
        data = pd.read_excel(i, sheet_name='PivotReqSrc')
        data = data.loc[:, list_1]
        New_Name = 'Sys-%s'%(time)
        print('importing 17/17.file\n')
        if os.path.exists(outpath_FRSys_5R1V + '%s.xlsx'%New_Name):
           os.remove(outpath_FRSys_5R1V + '%s.xlsx'%New_Name)
        data.to_excel(outpath_FRSys_5R1V+'%s.xlsx'%New_Name, index=False)


#os.system("pause")
