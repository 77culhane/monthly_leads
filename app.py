#!/usr/bin/env python
# coding: utf-8

# # Dependencies

# In[1]:


import pandas as pd
import datetime


# # Dependent Functions

# In[2]:


# function to get unique values 
def unique(list1): 
  
    # intilize a null list 
    unique_list = [] 
      
    # traverse for all elements 
    for x in list1: 
        # check if exists in unique_list or not 
        if x not in unique_list: 
            unique_list.append(x) 
    return unique_list

# quick add extra lines after the main df is built
def addline(rows, header, value):
    row = []
    for x in range(len(rows[1])):
        if x == 0:
            row.append(header)
        else:
            myvar = value.replace('alpha[x]',alpha[x])
            row.append(myvar)
    rows.append(row)
# get rid of 'nan' values; important for sorting
def clear_nans(mylist):
    for x in mylist:
        if str(x).lower() == "nan":
            mylist.remove(x)
        else:
            None


# In[3]:


#Import CSV; following initial generation of the output of this .py file, you can just cut and paste updated myreport.csv contents into the output .xlsx file
df = pd.read_csv('myreport.csv')


# In[4]:


#Salesforce exports .csv files with a text blurb at the end; this needs to be deleted
mylist = []
for x in range(7):
    mylist.append(len(df)-1-x)
df = df.drop(index=mylist)


# In[5]:


#Generate value lists; .xlsx formulas will use these values as COUNTIF criteria
progs_raw = unique(df["Anticipated Major"].to_list())
clear_nans(progs_raw)
progs = sorted(progs_raw)
statuses_raw = unique(df["Lead Status"].to_list())
clear_nans(statuses_raw)
statuses = sorted(statuses_raw)


# In[6]:


# Establish Timeline
yrs = ['2018','2019','2020','2021','2022','2023','2024','2025']
mns = ['01','02','03','04','05','06','07','08','09','10','11','12']
mydates = []
for y in yrs:
    for m in mns:
        mydt = f'{y},{m}'
        mydates.append(mydt)


# # Build Report

# In[7]:


#Main Section
#.xlsx file will need to refer to the contents of myreport.csv; store these contents under the following sheet name
sheetname = "DataBase"
#initialize list to contain the rows of the data table
rows = []
#prepare header
header = ["Degree"]+mydates
rows.append(header)
#this prevents errors
mydates.append(f'2026,1')
#populate table
for prog in progs:
    #first value of every row is the name of the degree
    row = [prog]
    #initialize an index counter
    x = 0
    #populate within range of mydates
    while x < len(mydates)-1:
        #line 24: range = Degree column, screen for each degree
        #line 25: range = Date column, screen for each date range
        #line 26: clean up the leftover " " strings; they break excel formulas
        cell = f"""=COUNTIFS(
        {sheetname}!B2:B20000,"{prog}",
        {sheetname}!A2:A20000,">="&DATE({mydates[x]},1),{sheetname}!A2:A20000,"<"&DATE({mydates[x+1]},1)
        )""".replace('\n','').replace("        ",'')
        #add cell to row
        row.append(cell)
        #counter
        x = x+1
    #add row to table
    rows.append(row)


# In[8]:


#This is how we will give the Excel Formulas the input they're expecting (i.e. CellA4 vs CellBB4)
alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
newalpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
for x in newalpha:
    for y in newalpha:
        z = x+y
        alpha.append(z)


# In[9]:


#Add Grand total at the bottom
#spacer
addline(rows, '','')
# Grand Totals
addline(rows, "Grand Total", f'=SUM(alpha[x]2:alpha[x]24)')


# In[10]:


df2 = pd.DataFrame(rows[1:])
df2.columns = rows[0]
df2

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('monthly_report.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name=sheetname, index=False)
df2.to_excel(writer, sheet_name='Monthly Leads', index=False)

# Close program and output finalreport.xlsx
writer.save()


# ### After the program is finished, user will still need to copy and paste the freshest copy of myreport.csv into the DataBase sheet of monthly_report.xlsx.
# ### The date values exported by Python are not recognized as "Dates" by Excel.
# ### Those same values cut and pasted from a report freshly exported from Salesforce ARE recognized.

# In[ ]:




