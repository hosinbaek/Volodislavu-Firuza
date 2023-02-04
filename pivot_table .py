#!/usr/bin/env python
# coding: utf-8

# In[1]:


#libraries
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sqlite3


# In[2]:


#driver options
options = Options()
options.add_argument("--headless")
driver = webdriver.Chrome(options=options)


# In[3]:


#load & download url
url = "https://jobs.homesteadstudio.co/wp-content/uploads/2022/12/skill_test_data.xlsx"
driver.get(url)


# In[4]:


#initialization of tabs
df1 = pd.read_excel("skill_test_data.xlsx", sheet_name="pivot_table")
df2 = pd.read_excel("skill_test_data.xlsx", sheet_name="data")


# In[5]:


#setting tab name
pivot_table=pd.pivot_table(df2,
    index =["Platform (Northbeam)"],
    values = ["Spend","Attributed Rev (1d)","Imprs","Visits","New Visits","Transactions (1d)","Email Signups (1d)"],
    aggfunc = "sum",
    margins = True
)
column_order = ["Spend","Attributed Rev (1d)","Imprs","Visits","New Visits","Transactions (1d)","Email Signups (1d)"]
pivot_table = pivot_table.reindex(column_order, axis = 1 )


# In[6]:


#renaming and sorting
pivot_table.index.name = 'Row Labels'
pivot_table.rename(columns={'Spend': 'Sum of Spend',
                            'Attributed Rev (1d)': 'Sum of Attributed Rev (1d)', 
                            'Imprs':'Sum of Imprs',
                            'Visits':'Sum of Visits',
                            'New Visits':'Sum of New Visits',
                            'Transactions (1d)':'Sum of Transactions (1d)',
                            'Email Signups (1d)': 'Sum of Email Signups (1d)',
                   }, inplace=True)
pivot_table.sort_values('Sum of Attributed Rev (1d)', ascending=False, inplace=True)
pivot_table.head()


# In[7]:


#create sqlite
conn = sqlite3.connect("pivot_table.db")
pivot_table.to_sql("pivot_table", conn, if_exists="replace")


# In[8]:


#close chrome driver
driver.quit()

