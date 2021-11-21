#!/usr/bin/python
# -*- coding: utf-8 -*-
import re, fitz, pdfplumber
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import xlsxwriter
#import openpyxl
from tkinter import *
from tkinter import filedialog
import glob,os,sys

class BreakOutNested(Exception):
   pass

def getdf(filename):
#Function for generating the policy's members details
    
# Reading the pdf with PyMuPDF library
 doc = fitz.open(filename)
 x = doc.metadata
 tfile = x['title']
 fname = tfile + '.txt'
 for page in doc:
    text = page.getText('text')
    sys.stdout = open(fname, 'a')
    print(text)
    sys.stdout.close()
 poldoc = open(fname)
 fpoldoc = poldoc.read()
 poldoc.close()
# Creation of the Regex objects for parsing the text files
 #Empid = re.compile(r'\s*Emp\s*ID\s*(.*)\s*(.*)\s*:')
 Empid = re.compile(r'\s*Emp\s*ID\s*([\s\S]*?):')
 Polno1 = re.compile(r'\s*Attached to and forming part of policy number\s*(.*)\s+')
 Polno2 = re.compile(r'Page\s+1\s+of\s+3\s+(.*)\s+')
 TPAid = re.compile(r'TPA\sID\s:\s(\w+)')
 TPAName = re.compile(r'TPA\s*Name\s*:\s*\n(.*)')
 Dependants1 = re.compile(r'Dependants\s(\d+)')
 Dependants2 = re.compile(r'Tel/Fax/Email\s*(\d+)')
 Issue_Office_Name = re.compile(r'\s*Issue Office Name\s+:\s+:\s+(.*)\s+(.*)')
 Product_Name1 = re.compile(r'(.*)\s*:Group Health Insurance Product')
 Product_Name2 = re.compile(r'(.*)\s*POLICY SCHEDULE')
 Policy_type1 = re.compile(r'\s*Mediclaim Insurance Policy\s*(.*)\s*POLICY')
 Policy_type2 = re.compile(r'POLICY-2017:\s*(.*?)\s*Policy')
 Coinsdetails = re.compile(r'Co-insurance Details\s*:\s*(.*)')
 prevpol = re.compile(r'(.*)\s*Prev.\s*Policy')
 from_to_date = re.compile(r'FROM 00:00  ON\s*(.*)\s*TO MIDNIGHT OF\s*(.*)\s*')
 Sum_Insured = re.compile(r'SI\s*(.*)\s*No Of\s*')
 '''In the next three patterns, notice that the raw string is not specified
 by preceding the expression with a r, as we have to specify the apostrophe 
 in the expression. The apostrophe is a character that has to be escaped by using a backslash'''
 Insured_Name1 = re.compile('Insured\'s Name\s*\n(.*)\n')
 Insured_Code = re.compile('Insured\'s Code\s*\n(.*)\n')
 Address1 = re.compile(r'Address\s*([\s\S]*?)12 DAYS')
 Address2 = re.compile(r'Address\s*([\s\S]*?)Address')
 Pincode = re.compile(r'\s*(\d+)\s*\nAddress')
 Dev_off_code = re.compile(r'\n(.*)\s*Dev.Off.Code')
 ''' The regex pattern below explains the difference between a greedy and non greedy match. * is greedy
 and ? is non-greedy. For e.g. if the input string on the next line after 'Agent/Broker Details' 
 is "BF0000001298 OBC RO  DELHI", then r'Agent/Broker Details\s*\n(.*) (.*)\s*'.group(1) is a greedy match and gives 
 output of 'BF0000001298 OBC RO' but r'Agent/Broker Details\s*\n(.*?) (.*)\s*'.group(1) is a non greedy match 
 and gives output of 'BF0000001298'. The greedy * goes up to the last space between RO and DELHI.
 But the non greedy ? goes up only till the space between BF0000001298 and OBC '''
 Agent_Name_Code = re.compile(r'Agent/Broker Details\s*\n(.*?) (.*)\s*')
 Gross_Premium = re.compile(r'\s*(.*)\s*\nGross Premium')
 Premium_data = re.compile(r'Gross Premium\s*\n\s*(.*)\s*\n\s*(.*)\s*\n\s*(.*)\s*')
 Collection_data = re.compile('CC (.*) - (.*)\s*GST')
 tel_email_data = re.compile(r'\s*/\s*/\s*(.*)/(.*)')
# Parsing the text files

# Regex search and assignment of TPAName
 if TPAName.search(fpoldoc):
    tpaname = TPAName.search(fpoldoc).group(1)
# Regex search and assignment of TPAid
 if TPAid.search(fpoldoc):
    tpaid = TPAid.search(fpoldoc).group(1)
# Regex search and assignment of Issue Office Name
 if Issue_Office_Name.search(fpoldoc):
    ion = Issue_Office_Name.search(fpoldoc).group(1) + Issue_Office_Name.search(fpoldoc).group(2)
# Regex search and assignment of Policy No and Issue Office Code
 if Polno1.search(fpoldoc):
    polno = Polno1.search(fpoldoc).group(1)
    ioc = polno[0:6]
 elif Polno2.search(fpoldoc):
    polno = Polno2.search(fpoldoc).group(1)
    ioc = polno[0:6]
 else:
    polno = None
    ioc = None
     # Regex search for Product Name and Plan Type; both have the same value
 if Product_Name1.search(fpoldoc):
    productname = Product_Name1.search(fpoldoc).group(1)
    plantype = productname
 elif Product_Name2.search(fpoldoc):
    productname = Product_Name2.search(fpoldoc).group(1)
    plantype = productname
     # Regex search for Policy type
 if Policy_type1.search(fpoldoc):
    poltype = Policy_type1.search(fpoldoc).group(1)
 elif Policy_type2.search(fpoldoc):
    poltype = Policy_type2.search(fpoldoc).group(1)
# Regex search and assignment of Number of Persons covered
 if Dependants1.search(fpoldoc):
    Dependants = Dependants1.search(fpoldoc).group(1)
    insured = int(Dependants) + 1
 elif Dependants2.search(fpoldoc):
    Dependants = Dependants2.search(fpoldoc).group(1)
    insured = int(Dependants)
 else:
    Dependants = 0
    insured = int(Dependants) + 1
# Regex search and assignment of Employee ID
 if Empid.search(fpoldoc):
    empid = Empid.search(fpoldoc).group(1)
    # Regex search for Co insurance details
 if Coinsdetails.search(fpoldoc):
    coinsdet = Coinsdetails.search(fpoldoc).group(1)
 else:
    coinsdet = 'NIL'
    # Regex search for Previous Policy
 if Sum_Insured.search(fpoldoc):
    SI = Sum_Insured.search(fpoldoc).group(1)
    sum_insured = SI
    # Regex search for From Date and To Date
    # Pl. note inception date and From_date are same
 if from_to_date.search(fpoldoc):
    from_date = from_to_date.search(fpoldoc).group(1)
    to_date = from_to_date.search(fpoldoc).group(2)
    inception_date = from_date
# Regex search for Sum Insured
 if prevpol.search(fpoldoc):
    prevpolicy = prevpol.search(fpoldoc).group(1)
#Regex search for Insured's Name
 if Insured_Name1.search(fpoldoc):
    Ins_Name = Insured_Name1.search(fpoldoc).group(1)
    sstring = "(GSTIN: 0)"
    if Ins_Name.endswith(sstring):
        Ins_Name = Ins_Name[:-10] # to remove the 10 character string (GSTIN :0) from the end
        emp_dependant_name = Ins_Name
    else:
        emp_dependant_name = Ins_Name
#Regex search for Insured's Code
 if Insured_Code.search(fpoldoc):
     insuredcode = Insured_Code.search(fpoldoc).group(1)
#Regex search for address
 if Address1.search(fpoldoc):
     address = Address1.search(fpoldoc).groups()
     address = ' '.join(address)
     address = address.replace('\n',' ')
     city = address
     state = address
 elif Address2.search(fpoldoc):
     address = Address2.search(fpoldoc).groups()
     address = ' '.join(address)
     address = address.replace('\n',' ')
     city = address
     state = address
#address = address.split('\n')
#address = ' '.join(str(x) for x in address)   
#address = ' '.join(str(x) for x in Address1).replace('\n',' ')
#Address1 = re.search('Address\s*([\s\S]*?)12', fpoldoc ).groups()
#address = ' '.join(str(x) for x in Address1).replace('\n',' ')
#stre = stre.replace('\n',' ')

#Regex search for Pincode
 if Pincode.search(fpoldoc):
     pincode = Pincode.search(fpoldoc).group(1)
 else:
     pincode = ' '
# Regex search for Dev_off_code
 if Dev_off_code.search(fpoldoc):
     dev_off_code = Dev_off_code.search(fpoldoc).group(1)
# Regex search for Agent Name
 if Agent_Name_Code.search(fpoldoc):
     agentcode = Agent_Name_Code.search(fpoldoc).group(1)
     agentname = Agent_Name_Code.search(fpoldoc).group(2)
# Regex search for Gross Premium
 if Gross_Premium.search(fpoldoc):
     grosspremium = Gross_Premium.search(fpoldoc).group(1)
# Regex search for ST,stamp duty and total premium
 if Premium_data.search(fpoldoc):
     service_tax = Premium_data.search(fpoldoc).group(1)
     stamp_duty = float(Premium_data.search(fpoldoc).group(2))
     total_premium = Premium_data.search(fpoldoc).group(3)
# Regex search for collection no and date
 if Collection_data.search(fpoldoc):
     collectionno = 'CC' + ' ' + Collection_data.search(fpoldoc).group(1)
     collectiondate = Collection_data.search(fpoldoc).group(2)
# Regex search for telephone and email
 if tel_email_data.search(fpoldoc):
     tel = tel_email_data.search(fpoldoc).group(1)
     email = tel_email_data.search(fpoldoc).group(2)
# Setting of fixed variables to NIL as these variables will always remain as NIL
 cover_note_no = 'NIL'
 cover_note_date = 'NIL'
 group_s_no = 'NIL'
 is_proposer_an_insured = 'NIL'
 marital_status = 'NIL'
 date_of_birth = 'NIL'
 cumulative_bonus= 'NIL'
 domicilliary_hospitalisation_limit= 'NIL'
 maternity_benefit_si= 'NIL'
 pre_existing_diseases = 'NIL'
 cumulative_bonus = 'NIL'
 domicilliary_hospitalisation_limit = 'NIL'
 maternity_benefit_si = 'NIL'
 pre_existing_diseases = 'NIL'
 dev_off_name = 'NIL'
 deductible = 'NIL'
 threshold_sum_insured = 'NIL'
 personal_accidental_cover_si = 'NIL'
 personal_accidental_cover_premium = 'NIL'
 basic_cover_si = 'NIL'
 basic_cover_premium = 'NIL'
 ncb_discount_si = 'NIL'
 ncb_discount_premium = 'NIL'
 daily_cash_allowance_si = 'NIL'
 daily_cash_allowance_premium = 'NIL'
 remarks = 'NIL'
 vb_compliance = 'NIL'
 enrollment = 'NIL'
 endorsement = 'NIL'
 endorsement_no = 'NIL'
 endorsement_effective_date = 'NIL'
 endorsement_type = 'NIL'
 endorsement_remarks = 'NIL'
 group_product = 'NIL'
# Getting the member data from the table on page 1 of the pdf
 pdf = pdfplumber.open(filename)
 page = pdf.pages[1]
 text = page.extract_text()

# For parsing the table for member details, the line below extracts
# all of the words between 'Any' and 'Total' including all whitespace characters
 m = re.search('Any\s*([\s\S]*?)Total', text)
 # Comparisons to singletons like None 
 # should always be done with is or is not, never the equality operators
 if m is None:
     sys.stdout = open("formaterror.txt", "a")
     print('this policy is in different format'+" "+ filename)
     sys.stdout.close()
     os.remove(fname)
     return None
 #if isinstance(m, None):
 #    print('m is none')
 k = m.groups()
# The line below creates a list of lists of all members from the table
 w = k[0].split('\n')
 w.pop()
 w.append('300')  # the 300 is a placeholder element to mark the end of the string
 temp = []
 finallist = []
 for x in w:
    if x[0].isdigit() == True and x[0] == w[0]:
        temp.append(x)
    elif x[0].isalpha() == True:
        temp.append(x)
    elif x[0].isdigit() == True and x[0] != w[0]:
        finallist.append(list(temp))
        temp.clear()
        temp.append(x)
 if w.index(x) == len(w) - 1:
    finallist.append(list(temp))

# using pop to perform removal of last element
 finallist.pop()
# using pop(0) to perform removal of first element of list
 finallist.pop(0)
# print(finallist)
 i = 0
 matches = ['Employed', 'Unemployed']
 while i <= len(finallist) - 1:
    mem_list = finallist[i]
    j = 0
    while j < len(mem_list):
        mem_list_ele = finallist[i][j]
        if j == 0:
            memsrnopatt = re.compile(r'(\d) (.*)')
            mem = re.search(memsrnopatt, mem_list_ele).group(1)
            namepatt = re.compile(r'(\d) (.*) (Self|Spouse|Dependant\sChild) (\w) (\d+) (\w+)')
            namepatt1 = re.compile(r'(\d) (.*) (Self|Spouse|Dependant\sChild) (\w) (\d+)')
            try:
                if re.search(namepatt, mem_list_ele) != None:
                    name1 = re.search(namepatt, mem_list_ele).group(2)
                    relationship1 = re.search(namepatt,mem_list_ele).group(3)
                    sex = re.search(namepatt, mem_list_ele).group(4)
                    age = re.search(namepatt, mem_list_ele).group(5)
                else:
                    name1 = re.search(namepatt1, mem_list_ele).group(2)
                    relationship1 = re.search(namepatt1,mem_list_ele).group(3)
                    sex = re.search(namepatt1, mem_list_ele).group(4)
                    age = re.search(namepatt1, mem_list_ele).group(5)
                    name2 = ' '
                    name3 = ' '
                    relationship2 = ' '
            except AttributeError:
                print('data not found')
     # ped = re.search(namepatt, mem_list_ele).group(6)
        '''if j == 1:
            if any(x in mem_list_ele for x in matches):
                k = re.compile(r'(.*)(Employed|Unemployed)')
                name2 = re.search(k, mem_list_ele).group(1)
                relationship2 = re.search(k, mem_list_ele).group(2)
                name3 = ' '
        if j == 1 and ' ' not in mem_list_ele:
            name2 = mem_list_ele
            name3 = ' '
            relationship2 = ' '
            '''
        if (j==1):
            if any(x in mem_list_ele for x in matches):
                k=re.compile(r'(.*)(Employed|Unemployed)\s*')
                name2 = re.search(k, mem_list_ele).group(1)
                relationship2 = re.search(k, mem_list_ele).group(2)
                name3 = ' '
            else:
                k=re.compile(r'(.*)')
                name2 = re.search(k,mem_list_ele).group(1)
                name2 = name2.strip()
                name3 =' '
                relationship2 = ' '

        if j == 2:
            name3 = mem_list_ele
        j += 1

  # Use the join method to concatenate multiple variables to a single variable
    name = ' '.join(str(x) for x in (name1, name2, name3))
    relationship = ' '.join(str(x) for x in (relationship1,relationship2))
  # print(mem)
  # print(name)

    df.loc[i] = [
        tpaid,tpaname,ioc,ion,coinsdet,polno,productname,poltype,
        plantype,cover_note_no,cover_note_date,prevpolicy,from_date,to_date,
        inception_date,insuredcode,insured,mem,group_s_no,empid,emp_dependant_name,
        SI,name,Ins_Name,is_proposer_an_insured,relationship,marital_status,date_of_birth,age,sex,
        sum_insured,cumulative_bonus,domicilliary_hospitalisation_limit,maternity_benefit_si, 
        pre_existing_diseases,address,city,state,pincode,tel,email,dev_off_code,
        dev_off_name,agentcode,agentname,grosspremium,service_tax,stamp_duty,total_premium,
        deductible,threshold_sum_insured,personal_accidental_cover_si,personal_accidental_cover_premium,
        basic_cover_si,basic_cover_premium,ncb_discount_si,ncb_discount_premium,daily_cash_allowance_si,
        daily_cash_allowance_premium,collectionno,collectiondate,remarks,vb_compliance,enrollment,endorsement,
        endorsement_no,endorsement_effective_date,endorsement_type,endorsement_remarks,group_product,
        ]
    i = i + 1
 return df

df = pd.DataFrame(columns=[
    'TPA_ID','TPA_NAME','ISSUE_OFFICE_CODE','ISSUE_OFFICE_NAME','CO_INSURANCE_DETAILS','POLICY_NO',
    'PRODUCT_NAME','POLICY_TYPE','PLAN_TYPE','COVER_NOTE_NO','COVER_NOTE_DATE','PREVIOUS_POLICY',
    'FROM_DATE','TO_DATE','INCEPTION_DATE','INSURED_CODE','No_of_persons_covered','MEMBER_S_NO',
    'GROUP_S_NO','EMP_ID','EMP/DEPENDANT NAME','SI','MEMBER_NAME','INSURED_NAME',
    'IS_PROPOSER_AN_INSURED','RELATIONSHIP','MARITAL_STATUS','DATE_OF_BIRTH','AGE','SEX',
    'SUM_INSURED','CUMULATIVE_BONUS','DOMICILLIARY_HOSPITALISATION_LIMIT','MATERNITY_BENEFIT_SI',
    'PRE_EXISTING_DISEASES','ADDRESS','CITY','STATE','PINCODE','TELEPHONE','EMAIL','DEV_OFF_CODE',
    'DEV_OFF_NAME','AGENT_CODE','AGENT_NAME','GROSS_PREMIUM','SERVICE_TAX','STAMP_DUTY',
    'TOTAL_PREMIUM','DEDUCTIBLE','THRESHOLD_SUM_INSURED','PERSONAL_ACCIDENTAL_COVER_SI',
    'PERSONAL_ACCIDENTAL_COVER_PREMIUM','BASIC_COVER_SI','BASIC_COVER_PREMIUM','NCB_DISCOUNT_SI',
    'NCB_DISCOUNT_PREMIUM','DAILY_CASH_ALLOWANCE_SI','DAILY_CASH_ALLOWANCE_PREMIUM','COLLECTION_NO',
    'COLLECTION_DATE','REMARKS','VB_COMPLIANCE','ENROLLMENT','ENDORSMENT','ENDORSMENT_NO',
    'ENDORSMENT_EFFECTIVE_DATE','ENDORSMENT_TYPE','ENDORSMENT_REMARKS','GROUP_PRODUCT'
    ])

writer = pd.ExcelWriter('OBCM.xlsx',engine='xlsxwriter')

folder_selected = filedialog.askdirectory()
files = [f for f in glob.glob(folder_selected + "**/*.pdf", recursive=True)]
os.chdir(folder_selected)
print('For errors, please look at error.txt')
ind = 0
for f in files:
    try:
        df=getdf(f)
        if df is None: raise BreakOutNested()
        else:
         if (ind == 0):
            df.to_excel(writer,index=False,startrow=ind)
            ind = ind + len(df) + 1
            #clearing the dataframe for populating the new policy no.
            df=df[0:0]
         else:
            df.to_excel(writer,index=False,startrow=ind,header=False)
            ind += len(df)
            #clearing the dataframe for populating the new policy no.
            df=df[0:0]

        #with pd.ExcelWriter('OBCM.xlsx',mode='a') as writer:
        #       df.to_excel(writer,sheet_name='Sheet2')
    except UnicodeEncodeError:
        sys.stdout = open("error.txt", "a")
        print('Error opening file'+" "+ f)
        sys.stdout.close()
    except BreakOutNested:
        df = pd.DataFrame(columns=[
            'TPA_ID','TPA_NAME','ISSUE_OFFICE_CODE','ISSUE_OFFICE_NAME','CO_INSURANCE_DETAILS','POLICY_NO',
            'PRODUCT_NAME','POLICY_TYPE','PLAN_TYPE','COVER_NOTE_NO','COVER_NOTE_DATE','PREVIOUS_POLICY',
            'FROM_DATE','TO_DATE','INCEPTION_DATE','INSURED_CODE','No_of_persons_covered','MEMBER_S_NO',
            'GROUP_S_NO','EMP_ID','EMP/DEPENDANT NAME','SI','MEMBER_NAME','INSURED_NAME',
            'IS_PROPOSER_AN_INSURED','RELATIONSHIP','MARITAL_STATUS','DATE_OF_BIRTH','AGE','SEX',
            'SUM_INSURED','CUMULATIVE_BONUS','DOMICILLIARY_HOSPITALISATION_LIMIT','MATERNITY_BENEFIT_SI',
            'PRE_EXISTING_DISEASES','ADDRESS','CITY','STATE','PINCODE','TELEPHONE','EMAIL','DEV_OFF_CODE',
            'DEV_OFF_NAME','AGENT_CODE','AGENT_NAME','GROSS_PREMIUM','SERVICE_TAX','STAMP_DUTY',
            'TOTAL_PREMIUM','DEDUCTIBLE','THRESHOLD_SUM_INSURED','PERSONAL_ACCIDENTAL_COVER_SI',
            'PERSONAL_ACCIDENTAL_COVER_PREMIUM','BASIC_COVER_SI','BASIC_COVER_PREMIUM','NCB_DISCOUNT_SI',
            'NCB_DISCOUNT_PREMIUM','DAILY_CASH_ALLOWANCE_SI','DAILY_CASH_ALLOWANCE_PREMIUM','COLLECTION_NO',
            'COLLECTION_DATE','REMARKS','VB_COMPLIANCE','ENROLLMENT','ENDORSMENT','ENDORSMENT_NO',
            'ENDORSMENT_EFFECTIVE_DATE','ENDORSMENT_TYPE','ENDORSMENT_REMARKS','GROUP_PRODUCT'
        ])

        pass
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
# Set the column width and format.
#worksheet.set_column('B:B', 22)
cell_format = workbook.add_format({'font_color': 'black'})
cell_format.set_align('center')
worksheet.set_column('A:BR', 25, cell_format)

writer.save()
'''
removing duplicates from an excel file based on values of columns POLICY_NO and MEMBER_NAME

df = pd.read_excel('OBCM.xlsx')
# keep=first as per drop_duplicates docs is default
df.drop_duplicates(subset=['MEMBER_NAME','POLICY_NO'],inplace=True) 

df.to_excel("memberlist.xlsx")'''
