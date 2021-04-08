#! /usr/bin/env python3

print("##########################################################################################\n\n")
print("                                      ITAM Script                                      ")
print("                   written by: Vikram Khosa (Panasonic Canada Inc)                     \n")
print("Description - The script takes raw ITAM reports from PNA (SQL) and PCI (Desktop Central) ")
print("as input and formats them according to the requirements provided by Japan ITAM team.")
print("\n\n##########################################################################################\n")
print("Formatting PNA ITAM data......")

import csv
import pandas as pd
import datetime
pd.options.mode.chained_assignment = None


fname = '01_HardwareInventoryRAW.csv'
fname2 = '02_SoftwareInventoryRAW.csv'

df=pd.read_csv(fname,encoding='UTF-8-sig',converters={'COMPANYID': lambda x: str(x)})
df2=pd.read_csv(fname2,encoding='UTF-8-sig')

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

#print(df.columns)

df2.rename(columns = {'ORGANIZATION':'COMPANY', 'Last_Logon_User1':'USER_LOGIN_NAME', 'Computer_Name':'PCID', 'Software_Name':'SOFTWARE_NAME', 'Software_Version':'SOFTWARE_VERSION', 'Installed_Location':'SOFTWARE_PATH', 'Software_Manufacturer':'SOFTWARE_VENDOR', 'Last_Asset_Scan_Time':'SCAN_DATE'}, inplace=True)
df2.insert(1,'PC_SITECODE','')
df2.insert(4,'PC_NAME','')


#Populating PC_NAME
df2['PC_NAME'] = df2['PCID']

#Removing duplicates
df2.drop_duplicates(subset=['PCID','SOFTWARE_NAME','SOFTWARE_VERSION'], inplace=True)

#Changing date format
df2['SCAN_DATE'] = pd.to_datetime(df2['SCAN_DATE']).dt.strftime('%Y/%m/%d %H:%M:%S')

#Removing devices from HardwareInventory that are not in SoftwareInventory and storing in new dataframe
df1 = df[df['DEVICE_ID'].isin(df2['PCID'])]

# print(df2.columns)
df2.sort_values('PCID', inplace=True) #Sorting csv file by PCID values

##Writing to csv and using line_terminator to convert CRLF to LF
#df1.to_csv('HVUD140RZZ_PNA',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 1(Regional Hardware Inventory Information)
#df2.to_csv('HVUD141RZZ_PNA',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 2(Regional Software Inventory Information)




################# REPORT 3 AND 4 ######################
fname3 = '03_MissingPatchesRAW.csv'
fname4 = '04_AntivirusReportRAW.csv'

df3=pd.read_csv(fname3,encoding='UTF-8-sig')
df4=pd.read_csv(fname4,encoding='UTF-8-sig')

#Renaming columns from the report
df3.rename(columns = {'Computer_Name':'PC_NO', 'KB_Number':'SOFT_CODE', 'Release_Date':'UPDATE_DATE'}, inplace=True)
df4.rename(columns = {'Computer_Name':'PC_NO', 'Software_Name':'SOFT_NAME', 'Software_Version':'SOFT_VERSION', 'Software_Installed_Date':'UPDATE_DATE'}, inplace=True)


#Removing devices that are not in SoftwareInventory and storing in new dataframe
df33 = df3[df3['PC_NO'].isin(df2['PCID'])]
df44 = df4[df4['PC_NO'].isin(df2['PCID'])]

df44.insert(1, 'SOFT_CODE', '')

#Removing duplicates
df33.drop_duplicates(subset=['PC_NO','SOFT_CODE','UPDATE_DATE'], inplace=True)
df44.drop_duplicates(subset=['PC_NO','SOFT_VERSION'], inplace=True)

#Splitting Software Version to major version only
###df44['SOFT_VERSION'] = df44['SOFT_VERSION'].astype(str)
###df44['SOFT_CODE'] = df44['SOFT_VERSION'].str.split('.').str[0]

#Changing Software Name to match the formart "Pattern File Ver.14"
df44['SOFT_NAME'] = 'Pattern File Ver.' + df44['SOFT_VERSION']

#Changing Software Code to match the format "#AV_PTN_V14"
df44['SOFT_CODE'] = '#AV_PTN_V' + df44['SOFT_VERSION']

#Replace dot with underscore in SOFT_CODE for AV report
df44['SOFT_CODE'] = df44['SOFT_CODE'].str.replace('.', '_')

#Changing date format
df44['UPDATE_DATE'] = pd.to_datetime(df44['UPDATE_DATE']).dt.strftime('%Y%m%d%H%M%S')
df44['UPDATE_DATE'] = pd.to_datetime('today').strftime('%Y%m%d%H%M%S') #changing date to report generated date


#Removing entries which dont match the KB number pattern
df33['NEWcolumn'] = df33['SOFT_CODE'].str.contains(r'[\d]{7}',na=False)
df33 = df33[df33.NEWcolumn != False]

#Removing unwanted columns
df33.drop(['Patch_ID', 'Bulletin_ID', 'Patch_Description', 'Patch_Status', 'Vendor', 'NEWcolumn'], axis=1, inplace=True)

#Setting date on columns which have "--" as date
###d = datetime.datetime(1969, 12, 31, 20, 00)
###df33.loc[df33['UPDATE_DATE'] == '--', 'UPDATE_DATE'] = d

#Changing date format
df33['UPDATE_DATE'] = pd.to_datetime(df33['UPDATE_DATE']).dt.strftime('%Y%m%d%H%M%S')
df33['UPDATE_DATE'] = pd.to_datetime('today').strftime('%Y%m%d%H%M%S') #changing date to report generated date

#################### Creating Report 4 ####################
df5 = df33.copy()
df5.drop(['PC_NO'], axis=1, inplace=True)
df5.insert(1,'SOFT_NAME','')
df5.insert(3,'POLICY_CLASS','')
df5.insert(4,'SOFT_VERSION','')

#Changing SOFTWARE_NAME and SOFT_CODE format to match the requirements
df5['SOFT_NAME'] = 'KB' + df5['SOFT_CODE'] + ' not applied'
df5['SOFT_CODE'] = '#MSPU_KB' + df5['SOFT_CODE']

#Creating a copy of df44 to clean the dataframe and sort it by AV version.
df444 = df44.copy()
df444.drop(['PC_NO'], axis=1, inplace=True)
df444.drop_duplicates(subset=['SOFT_CODE', 'SOFT_NAME', 'SOFT_VERSION'], inplace=True)
df444.sort_values('SOFT_CODE', ascending=False, inplace=True)
df444['SOFT_VERSION'] = ''  #SOFT_VERSION is not required

#Adding top 10 entries of antivirus data from df444 to df5 and then removing duplicates
df5 = df5.append(df444.head(10))
df5.drop_duplicates(subset=['SOFT_CODE', 'SOFT_NAME', 'SOFT_VERSION'], inplace=True)

#################### End of Report 4 ########################

#Changing SOFT_CODE format to match the requirements - df['col'] = 'str' + df['col'].astype(str)
df33['SOFT_CODE'] = '#MSPU_KB' + df33['SOFT_CODE']

df33.insert(3,'PRODUCT_ID','') #Adding columns not in the report

#Adding antivirus data from df44 to df33 and then removing SOFTWARE_NAME and SOFT_VERSION
df33 = df33.append(df44)
df33.drop(['SOFT_NAME','SOFT_VERSION'], axis=1, inplace=True)

df33.sort_values('PC_NO', inplace=True) #Sorting csv file by PC_NO

## df44.to_csv('04_AntiVirusPNA.csv',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Not required to print
#df33.to_csv('HVUD142RZZ_PNA',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 3(Regional Software Inventory Information)
#df5.to_csv('HVUD143RZZ_PNA',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 4(Regional Software License Information)


#####################################################################
#################### Running PCI ITAM Script ########################
#####################################################################

print("Formatting PCI ITAM data.....")


fname = '01_HardwareInventoryPCIRAW.csv'
fname2 = '02_SoftwareInventoryPCIRAW.csv'

dfpci=pd.read_csv(fname,encoding='ISO-8859-1',converters={'COMPANYID': lambda x: str(x)})
df2pci=pd.read_csv(fname2,encoding='ISO-8859-1')

#Renaming columns from the report
dfpci.rename(columns = {'computer_name':'DEVICE_ID', 'organization_name':'ORGANIZATION', 'OWNER':'PERSON', 'operating_system':'OS', 'service_pack':'OS_LEVEL', 'family':'CPU', 'clock_speed':'CPU_SPEED', 'bios_version':'BIOS', 'total_physical_memory_mb':'MEMORY', 'ip_address':'IPADDR', 'ip_subnet':'SUBNETMASK', 'mac_address':'MACADDR', 'total_capacity':'SYSDRV_TOTAL', 'free_space':'SYSDRV_FREE', 'owner_email_id':'OWNERID', 'domain_name':'DOMAIN', 'manufacturer':'WMANUFACTURER', 'device_model':'WMODEL', 'no_of_processors':'WNUMBEROFPROCESSORS', 'servicetag':'SERIAL_NUMBER', 'system_type':'HWARCHITECTURE', 'build_number':'OS_BUILDNO', 'LAST_SUCCESSFUL_SCAN':'LASTCONNECTDATE'}, inplace=True)

#Adding columns not in the report
dfpci.insert(1,'DEVICENAME','')
dfpci.insert(9,'KEYBOARD','')
dfpci.insert(10,'MOUSE','')
dfpci.insert(11,'VIDEO','')
dfpci.insert(13,'LANGUAGE','')
dfpci.insert(16,'SUBNETADDR','')
dfpci.insert(18,'N_PARALLEL','')
dfpci.insert(19,'N_SERIAL','')
dfpci.insert(20,'N_PRINTER','')
dfpci.insert(21,'SYSDRV','')
dfpci.insert(24,'CHASSISTYPE','')
dfpci.insert(25,'PCID','')
dfpci.insert(26,'SITECODE','')
dfpci.insert(28,'REGION','')
dfpci.insert(31,'PILOT','')
dfpci.insert(32,'RESPONSIBLEIT','')
dfpci.insert(33,'WDOMAIN','')
dfpci.insert(37,'WSYSTEMTYPE','')
dfpci.insert(38,'WVIDEOMODEDESCRIPTION','')
dfpci.insert(40,'DESKTOP_OR_LAPTOP','')
dfpci.insert(41,'WDNSHOSTNAME','')
dfpci.insert(42,'WDNSDOMAIN','')
dfpci.insert(43,'WDNSDOMAINSUFFIXSEARCHORDER','')
dfpci.insert(45,'SWADDRESSWIDTH','')
dfpci.insert(48,'ADMINISTRATORID','')
dfpci.insert(49,'INSTALLATIONBUILDING','')
dfpci.insert(50,'FIXEDASSETNUMBER','')

#Populating DEVICENAME and FIXEDASSETNUMBER
dfpci['DEVICENAME'] = dfpci['DEVICE_ID']
dfpci['FIXEDASSETNUMBER'] = dfpci['DEVICE_ID']

#Converting disk space from GB to MB
dfpci['SYSDRV_TOTAL'] =dfpci['SYSDRV_TOTAL']*1024
dfpci['SYSDRV_FREE'] =dfpci['SYSDRV_FREE']*1024

#Removing duplicates
dfpci.drop_duplicates(subset=['DEVICE_ID'], inplace=True)

#Remove computers not in PCI domain
#dfpci = dfpci[dfpci.DOMAIN != 'WORKGROUP']

#Splitting IP Address to single IP
dfpci['IPADDR'] = dfpci['IPADDR'].str.split(',').str[0]

#Removing SUBNETMASK suffix
dfpci['SUBNETMASK'] = dfpci['SUBNETMASK'].str.split(',').str[0]

#Changing MAC Address format by removing colon
dfpci['MACADDR'] = dfpci['MACADDR'].str.replace(':', '')

#Splitting OS_BUILDNO to major version only
dfpci['OS_BUILDNO'] = dfpci['OS_BUILDNO'].astype(str)
dfpci['OS_BUILDNO'] = dfpci['OS_BUILDNO'].str.split('.').str[0]

#Changing date format
dfpci['LASTCONNECTDATE'] = pd.to_datetime(dfpci['LASTCONNECTDATE']).dt.strftime('%Y/%m/%d %H:%M:%S')

#print(dfpci.columns)

df2pci.rename(columns = {'ORGANIZATION':'COMPANY', 'Last Logon User':'USER_LOGIN_NAME', 'Computer Name':'PCID', 'Software Name':'SOFTWARE_NAME', 'Software Version':'SOFTWARE_VERSION', 'Installed Location':'SOFTWARE_PATH', 'Software Manufacturer':'SOFTWARE_VENDOR', 'Last Asset Scan Time':'SCAN_DATE'}, inplace=True)
df2pci.insert(1,'PC_SITECODE','')
df2pci.insert(4,'PC_NAME','')

#Populating PC_NAME
df2pci['PC_NAME'] = df2pci['PCID']

#Removing duplicates
df2pci.drop_duplicates(subset=['PCID','SOFTWARE_NAME','SOFTWARE_VERSION'], inplace=True)

#Dropping invalid rows
df2pci = df2pci.dropna(axis=0, subset=['SOFTWARE_NAME'])

#Changing date format
df2pci['SCAN_DATE'] = pd.to_datetime(df2pci['SCAN_DATE']).dt.strftime('%Y/%m/%d %H:%M:%S')

#Removing devices from HardwareInventory that are not in SoftwareInventory and vice versa and storing in new dataframe
df1pci = dfpci[dfpci['DEVICE_ID'].isin(df2pci['PCID'])]
df22pci = df2pci[df2pci['PCID'].isin(df1pci['DEVICE_ID'])]

#print(df2pci.columns)
df22pci.sort_values('PCID', inplace=True) #Sorting csv file by PCID values

##Writing to csv and using line_terminator to convert CRLF to LF
#df1pci.to_csv('HVUD140RZZ_PCI',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 1
#df2pci.to_csv('HVUD141RZZ_PCI',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 2




################# REPORT 3 AND 4 ######################
fname3 = 'Detailed Patch Summary.csv'
fname4 = '04_AntivirusReportPCIRAW.csv'

df3pci=pd.read_csv(fname3,encoding='ISO-8859-1')
df4pci=pd.read_csv(fname4,encoding='ISO-8859-1')

#Renaming columns from the report
df3pci.rename(columns = {'Computer Name':'PC_NO', 'KB Number':'SOFT_CODE', 'Release Date':'UPDATE_DATE'}, inplace=True)
df4pci.rename(columns = {'Computer Name':'PC_NO', 'Software Name':'SOFT_NAME', 'Software Version':'SOFT_VERSION', 'Software Installed Date':'UPDATE_DATE'}, inplace=True)


#Removing devices that are not in SoftwareInventory and storing in new dataframe
df33pci = df3pci[df3pci['PC_NO'].isin(df22pci['PCID'])]
df44pci = df4pci[df4pci['PC_NO'].isin(df22pci['PCID'])]

df44pci.insert(1, 'SOFT_CODE', '')

#Removing duplicates
df33pci.drop_duplicates(subset=['PC_NO','SOFT_CODE','UPDATE_DATE'], inplace=True)
df44pci.drop_duplicates(subset=['PC_NO','SOFT_VERSION'], inplace=True)

#Splitting Software Version to major version only
#df44pci['SOFT_VERSION'] = df44pci['SOFT_VERSION'].astype(str)
#df44pci['SOFT_CODE'] = df44pci['SOFT_VERSION'].str.split('.').str[0]

#Changing Software Name to match the formart "Pattern File Ver.14"
df44pci['SOFT_NAME'] = 'Pattern File Ver.' + df44pci['SOFT_VERSION']

#Changing Software Code to match the format "#AV_PTN_V14"
df44pci['SOFT_CODE'] = '#AV_PTN_V' + df44pci['SOFT_VERSION']

#Replace dot with underscore in SOFT_CODE for AV report
df44pci['SOFT_CODE'] = df44pci['SOFT_CODE'].str.replace('.', '_')

#Changing date format
df44pci['UPDATE_DATE'] = pd.to_datetime(df44pci['UPDATE_DATE']).dt.strftime('%Y%m%d%H%M%S')
df44pci['UPDATE_DATE'] = pd.to_datetime('today').strftime('%Y%m%d%H%M%S') #changing date to report generated date


#Removing entries which dont match the KB number pattern
df33pci['NEWcolumn'] = df33pci['SOFT_CODE'].str.contains(r'[\d]{7}',na=False)
df33pci = df33pci[df33pci.NEWcolumn != False]

#Removing unwanted columns
df33pci.drop(['Patch ID', 'Bulletin ID', 'Patch Description', 'Patch Status', 'Vendor', 'NEWcolumn'], axis=1, inplace=True)

#Setting date on columns which have "--" as date
#d = datetime.datetime(1969, 12, 31, 20, 00)
#df33pci.loc[df33pci['UPDATE_DATE'] == '--', 'UPDATE_DATE'] = d

#Changing date format
df33pci['UPDATE_DATE'] = pd.to_datetime(df33pci['UPDATE_DATE']).dt.strftime('%Y%m%d%H%M%S')
df33pci['UPDATE_DATE'] = pd.to_datetime('today').strftime('%Y%m%d%H%M%S') #changing date to report generated date

#################### Creating Report 4 ####################
df5pci = df33pci.copy()
df5pci.drop(['PC_NO'], axis=1, inplace=True)
df5pci.insert(1,'SOFT_NAME','')
df5pci.insert(3,'POLICY_CLASS','')
df5pci.insert(4,'SOFT_VERSION','')

#Changing SOFTWARE_NAME and SOFT_CODE format to match the requirements
df5pci['SOFT_NAME'] = 'KB' + df5pci['SOFT_CODE'] + ' not applied'
df5pci['SOFT_CODE'] = '#MSPU_KB' + df5pci['SOFT_CODE']

#Creating a copy of df44pci to clean the dataframe and sort it by AV version.
df444pci = df44pci.copy()
df444pci.drop(['PC_NO'], axis=1, inplace=True)
df444pci.drop_duplicates(subset=['SOFT_CODE', 'SOFT_NAME', 'SOFT_VERSION'], inplace=True)
df444pci.sort_values('SOFT_CODE', ascending=False, inplace=True)
df444pci['SOFT_VERSION'] = ''  #SOFT_VERSION is not required

#Adding top 10 entries of antivirus data from df444pci to df5pci and then removing duplicates
df5pci = df5pci.append(df444pci.head(10))
df5pci.drop_duplicates(subset=['SOFT_CODE', 'SOFT_NAME', 'SOFT_VERSION'], inplace=True)

#################### End of Report 4 ########################

#Changing SOFT_CODE format to match the requirements - dfpci['col'] = 'str' + dfpci['col'].astype(str)
df33pci['SOFT_CODE'] = '#MSPU_KB' + df33pci['SOFT_CODE']

df33pci.insert(3,'PRODUCT_ID','') #Adding columns not in the report

#Adding antivirus data from df44pci to df33pci and then removing SOFTWARE_NAME and SOFT_VERSION
df33pci = df33pci.append(df44pci)
df33pci.drop(['SOFT_NAME','SOFT_VERSION'], axis=1, inplace=True)

df33pci.sort_values('PC_NO', inplace=True) #Sorting csv file by PC_NO

##df44pci.to_csv('04_AntiVirusPCI.csv',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n')
#df33pci.to_csv('HVUD142RZZ_PCI',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 3
#df5pci.to_csv('HVUD143RZZ_PCI',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 4

df1_combined = pd.concat([df1,df1pci])
df2_combined = pd.concat([df2,df22pci])
df33_combined = pd.concat([df33,df33pci])
df5_combined = pd.concat([df5,df5pci])

df1_combined.to_csv('HVUD140RZZ',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 1
df2_combined.to_csv('HVUD141RZZ',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 2
df33_combined.to_csv('HVUD142RZZ',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 3
df5_combined.to_csv('HVUD143RZZ',index=False,quoting=csv.QUOTE_ALL,line_terminator='\n') #Saving report 4

print("Reports Generated Successfully!!")
