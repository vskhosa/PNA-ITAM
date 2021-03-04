#! usr/bin/env python3

import csv
import pandas as pd
import datetime

fname = '01_HardwareInventoryRAW.csv'
fname2 = '02_SoftwareInventoryRAW.csv'

df=pd.read_csv(fname,encoding='UTF-8-sig',converters={'COMPANYID': lambda x: str(x)})
df2=pd.read_csv(fname2,encoding='UTF-8-sig')

print(df.columns)
print(df2.columns)

#Renaming columns from the report
df.rename(columns = {'computer_name':'DEVICE_ID', 'organization_name':'ORGANIZATION', 'OWNER':'PERSON', 'operating_system':'OS', 'service_pack':'OS_LEVEL', 'family':'CPU', 'clock_speed':'CPU_SPEED', 'bios_version':'BIOS', 'total_physical_memory_mb':'MEMORY', 'ip_address':'IPADDR', 'ip_subnet':'SUBNETMASK', 'mac_address':'MACADDR', 'total_capacity':'SYSDRV_TOTAL', 'free_space':'SYSDRV_FREE', 'owner_email_id':'OWNERID', 'domain_name':'DOMAIN', 'manufacturer':'WMANUFACTURER', 'device_model':'WMODEL', 'no_of_processors':'WNUMBEROFPROCESSORS', 'servicetag':'SERIAL_NUMBER', 'system_type':'HWARCHITECTURE', 'build_number':'OS_BUILDNO', 'LAST_SUCCESSFUL_SCAN':'LASTCONNECTDATE'}, inplace=True)

#Adding columns not in the report
df.insert(1,'DEVICENAME','')
df.insert(9,'KEYBOARD','')
df.insert(10,'MOUSE','')
df.insert(11,'VIDEO','')
df.insert(13,'LANGUAGE','')
df.insert(16,'SUBNETADDR','')
df.insert(18,'N_PARALLEL','')
df.insert(19,'N_SERIAL','')
df.insert(20,'N_PRINTER','')
df.insert(21,'SYSDRV','')
df.insert(24,'CHASSISTYPE','')
df.insert(25,'PCID','')
df.insert(26,'SITECODE','')
df.insert(28,'REGION','')
df.insert(31,'PILOT','')
df.insert(32,'RESPONSIBLEIT','')
df.insert(33,'WDOMAIN','')
df.insert(37,'WSYSTEMTYPE','')
df.insert(38,'WVIDEOMODEDESCRIPTION','')
df.insert(40,'DESKTOP_OR_LAPTOP','')
df.insert(41,'WDNSHOSTNAME','')
df.insert(42,'WDNSDOMAIN','')
df.insert(43,'WDNSDOMAINSUFFIXSEARCHORDER','')
df.insert(45,'SWADDRESSWIDTH','')
df.insert(48,'ADMINISTRATORID','')
df.insert(49,'INSTALLATIONBUILDING','')
df.insert(50,'FIXEDASSETNUMBER','')


#Populating DEVICENAME and FIXEDASSETNUMBER
df['DEVICENAME'] = df['DEVICE_ID']
df['FIXEDASSETNUMBER'] = df['DEVICE_ID']

#Converting disk space from GB to MB
df['SYSDRV_TOTAL'] =df['SYSDRV_TOTAL']*1024
df['SYSDRV_FREE'] =df['SYSDRV_FREE']*1024

#Removing duplicates
df.drop_duplicates(subset=['DEVICE_ID'], inplace=True)

#Remove computers not in PCI domain
#df = df[df.DOMAIN != 'WORKGROUP']

#Splitting IP Address to single IP
df['IPADDR'] = df['IPADDR'].str.split(',').str[0]

#Removing SUBNETMASK suffix
df['SUBNETMASK'] = df['SUBNETMASK'].str.split(',').str[0]

#Changing MAC Address format by removing colon
df['MACADDR'] = df['MACADDR'].str.replace(':', '')

#Splitting OS_BUILDNO to major version only
df['OS_BUILDNO'] = df['OS_BUILDNO'].astype(str)
df['OS_BUILDNO'] = df['OS_BUILDNO'].str.split('.').str[0]

#Changing date format
df['LASTCONNECTDATE'] = pd.to_datetime(df['LASTCONNECTDATE']).dt.strftime('%Y/%m/%d %H:%M:%S')

print(df.columns)

df2.rename(columns = {'ORGANIZATION':'COMPANY', 'Last_Logon_User1':'USER_LOGIN_NAME', 'Computer_Name':'PCID', 'Software_Name':'SOFTWARE_NAME', 'Software_Version':'SOFTWARE_VERSION', 'Installed_Location':'SOFTWARE_PATH', 'Software_Manufacturer':'SOFTWARE_VENDOR', 'Last_Asset_Scan_Time':'SCAN_DATE'}, inplace=True)
df2.insert(1,'PC_SITECODE','')
df2.insert(4,'PC_NAME','')

print(df2.columns)

#Populating PC_NAME
df2['PC_NAME'] = df2['PCID']

#Removing duplicates
df2.drop_duplicates(subset=['PCID','SOFTWARE_NAME','SOFTWARE_VERSION'], inplace=True)

#Changing date format
df2['SCAN_DATE'] = pd.to_datetime(df2['SCAN_DATE']).dt.strftime('%Y/%m/%d %H:%M:%S')

#Removing devices from HardwareInventory that are not in SoftwareInventory and storing in new dataframe
df1 = df[df['DEVICE_ID'].isin(df2['PCID'])]

print(df2.columns)
df2.sort_values('PCID', inplace=True) #Sorting csv file by PCID values

#Writing to csv and using line_terminator to convert CRLF to LF
df1.to_csv('01_HardwareInventoryPCI.csv',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 1
df2.to_csv('02_SoftwareInventoryPCI.csv',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 2




################# REPORT 3 AND 4 ######################
#fname3 = '03_MissingPatchesRAW.csv'
#fname4 = '04_AntivirusReportRAW.csv'

#df3=pd.read_csv(fname3,encoding='ISO-8859-1')
#df4=pd.read_csv(fname4,encoding='ISO-8859-1')

#Renaming columns from the report
#df3.rename(columns = {'Computer Name':'PC_NO', 'KB Number':'SOFT_CODE', 'Release Date':'UPDATE_DATE'}, inplace=True)
#df4.rename(columns = {'Computer Name':'PC_NO', 'Software Name':'SOFT_NAME', 'Software Version':'SOFT_VERSION', 'Software Installed Date':'UPDATE_DATE'}, inplace=True)


#Removing devices that are not in SoftwareInventory and storing in new dataframe
#df33 = df3[df3['PC_NO'].isin(df2['PCID'])]
#df44 = df4[df4['PC_NO'].isin(df2['PCID'])]

#df44.insert(1, 'SOFT_CODE', '')

#Removing duplicates
#df33.drop_duplicates(subset=['PC_NO','SOFT_CODE','UPDATE_DATE'], inplace=True)
#df44.drop_duplicates(subset=['PC_NO','SOFT_VERSION'], inplace=True)

#Splitting Software Version to major version only
###df44['SOFT_VERSION'] = df44['SOFT_VERSION'].astype(str)
###df44['SOFT_CODE'] = df44['SOFT_VERSION'].str.split('.').str[0]

#Changing Software Name to match the formart "Pattern File Ver.14"
#df44['SOFT_NAME'] = 'Pattern File Ver.' + df44['SOFT_VERSION']

#Changing Software Code to match the format "#AV_PTN_V14"
#df44['SOFT_CODE'] = '#AV_PTN_V' + df44['SOFT_VERSION']

#Replace dot with underscore in SOFT_CODE for AV report
#df44['SOFT_CODE'] = df44['SOFT_CODE'].str.replace('.', '_')

#Changing date format
#df44['UPDATE_DATE'] = pd.to_datetime(df44['UPDATE_DATE']).dt.strftime('%Y%m%d%H%M%S')
#df44['UPDATE_DATE'] = pd.to_datetime('today').strftime('%Y%m%d%H%M%S') #changing date to report generated date


#Removing entries which dont match the KB number pattern
#df33['NEWcolumn'] = df33['SOFT_CODE'].str.contains(r'[\d]{7}',na=False)
#df33 = df33[df33.NEWcolumn != False]

#Removing unwanted columns
#df33.drop(['Patch ID', 'Bulletin ID', 'Patch Description', 'Patch Status', 'Vendor', 'NEWcolumn'], axis=1, inplace=True)

#Setting date on columns which have "--" as date
###d = datetime.datetime(1969, 12, 31, 20, 00)
###df33.loc[df33['UPDATE_DATE'] == '--', 'UPDATE_DATE'] = d

#Changing date format
#df33['UPDATE_DATE'] = pd.to_datetime(df33['UPDATE_DATE']).dt.strftime('%Y%m%d%H%M%S')
#df33['UPDATE_DATE'] = pd.to_datetime('today').strftime('%Y%m%d%H%M%S') #changing date to report generated date

#################### Creating Report 4 ####################
#df5 = df33.copy()
#df5.drop(['PC_NO'], axis=1, inplace=True)
#df5.insert(1,'SOFT_NAME','')
#df5.insert(3,'POLICY_CLASS','')
#df5.insert(4,'SOFT_VERSION','')

#Changing SOFTWARE_NAME and SOFT_CODE format to match the requirements
#df5['SOFT_NAME'] = 'KB' + df5['SOFT_CODE'] + ' not applied'
#df5['SOFT_CODE'] = '#MSPU_KB' + df5['SOFT_CODE']

#Creating a copy of df44 to clean the dataframe and sort it by AV version.
#df444 = df44.copy()
#df444.drop(['PC_NO'], axis=1, inplace=True)
#df444.drop_duplicates(subset=['SOFT_CODE', 'SOFT_NAME', 'SOFT_VERSION'], inplace=True)
#df444.sort_values('SOFT_CODE', ascending=False, inplace=True)
#df444['SOFT_VERSION'] = ''  #SOFT_VERSION is not required

#Adding top 10 entries of antivirus data from df444 to df5 and then removing duplicates
#df5 = df5.append(df444.head(10))
#df5.drop_duplicates(subset=['SOFT_CODE', 'SOFT_NAME', 'SOFT_VERSION'], inplace=True)

#################### End of Report 4 ########################

#Changing SOFT_CODE format to match the requirements - df['col'] = 'str' + df['col'].astype(str)
#df33['SOFT_CODE'] = '#MSPU_KB' + df33['SOFT_CODE']

#df33.insert(3,'PRODUCT_ID','') #Adding columns not in the report

#Adding antivirus data from df44 to df33 and then removing SOFTWARE_NAME and SOFT_VERSION
#df33 = df33.append(df44)
#df33.drop(['SOFT_NAME','SOFT_VERSION'], axis=1, inplace=True)

#df33.sort_values('PC_NO', inplace=True) #Sorting csv file by PC_NO

#df44.to_csv('04_AntiVirusPCI.csv',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n')
#df33.to_csv('HVUD142RZZ',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 3
#df5.to_csv('HVUD143RZZ',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 4
