#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().run_line_magic('run', 'functions.py')
# BOM
machined, purchased = loadBOM('Data/GF12 BOM Master Tab Copies 1.27.24.xlsx')
#PO
po = loadPO('Data/GF12 All Purchase Orders By Job 1.27.24.xlsx')
po = mapPOtoPurchased(po, purchased)


print('PO')
display(po.head())
print('BOM Machined')
display(machined.head())
print('BOM Purchased')
display(purchased.head())


# In[2]:


# print duplicates in po['Part Number'] that have different Override 1 values
print('Duplicates in PO')
dup = po[po.duplicated(subset=['Part Number'], keep=False)].sort_values(by=['Part Number'])
uniqueParts = dup['Part Number'].unique()
display(dup[dup['Part Number'] == uniqueParts[4]])


# In[3]:


lookupPartNumber(po, '001283')


# In[28]:


get_ipython().run_line_magic('run', 'functions.py')
final = process_parts(po, machined, purchased, verbose=True)


# In[5]:


final

