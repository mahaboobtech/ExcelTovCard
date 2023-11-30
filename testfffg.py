import xlrd
import pandas as pd
import os
from datetime import datetime

#file=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop','Vcard','Contacts.xlsx') #If your excel file not in same directory with python file u can use it 
file= 'Contacts1.xlsx' #If your excel file in same directory with python file u can use it 
excelfile= pd.ExcelFile(file)
column = excelfile.parse('Workers')
s = ""
begin = "BEGIN:VCARD\nVERSION:2.1"

for i in range(len(column)):
    fName=""
    sName=""
    mName= ""
    prefix =""
    suffix =""

    secMail=""
    secOrg=""
    secTit=""
    BirDay=""
    if(str(column["Phone"][i])!="nan"):
        if(str(column["Name"][i])!="nan"):
            fName=str(column["Name"][i])
        if(str(column["Surname"][i])!="nan"):
            sName=str(column["Surname"][i])
        if("Suffix" in column.columns.values):
            suffix = ""
            if(str(column["Suffix"][i])!="nan"):
                suffix=str(column["Suffix"][i])+" "
        if("Prefix" in column.columns.values):
            prefix=""
            if(str(column["Prefix"][i])!="nan"):
                prefix=str(column["Prefix"][i])+" "
                
        if("MiddleName" in column.columns.values):
            mName=" "
            if(str(column["MiddleName"][i])!="nan"):
                mName=" " + str(column["MiddleName"][i])+" "
        secN="\nN:"+ sName + ";" + fName + ";"+mName.strip()+";"+prefix.strip()+";"+suffix.strip()
        secFN="\nFN:" + prefix + "" + fName + mName + sName + "," + suffix
        secPhone="\nTEL;CELL:+91"+str(column["Phone"][i]).split(".")[0] #v1
        # secPhone="\nTEL;CELL:+91"+str(column["Phone"][i]) #v2
        # print("Phone ",secPhone) #For testing purposes
        if("Mail" in column.columns.values):
            secMail=""
            if(str(column["Mail"][i]) != "nan"):
                secMail="\nEMAIL;HOME:"+str(column["Mail"][i])
        if("Organization" in column.columns.values):
            secOrg=""
            if(str(column["Organization"][i]) != "nan"):
                secOrg="\nORG:" + str(column["Organization"][i])
        if("Title" in column.columns.values):
            secTit=""
            if(str(column["Title"][i]) != "nan"):
                secTit="\nTITLE:" + str(column["Title"][i])
        if("DateOfBirth" in column.columns.values):
            BirDay=""
            if(str(column["DateOfBirth"][i]) != "nan"):
                varstr=str(column["DateOfBirth"][i])
                dt = datetime.strptime(varstr, '%Y-%m-%d %H:%M:%S')         
                BirDay="\nBDAY:" + str(dt.year) + str("-")+str(dt.month) + str("-") + str(dt.day)
        s+=begin+secN + secFN +secPhone + secMail+ secOrg+ secTit+BirDay+"\nEND:VCARD\n"
text_file = open("Exported.vcf", "w",encoding="utf-8") #Encoding utf-8 added
text_file.write(s)
text_file.close()
print("Completed!")
