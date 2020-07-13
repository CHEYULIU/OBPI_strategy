#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import datetime
from pandas import DataFrame, Series
 
unix_ts = 1439111214.0
time = datetime.datetime.fromtimestamp(unix_ts)


# In[6]:


f = lambda x: datetime.datetime.fromtimestamp(x).strftime("%Y-%m-%d")
def eu_call(s,k,T,t,sig,rf):
    d1 = (log(s/k)+(rf+0.5*sig**2)*(T-t))/(sig*sqrt(T-t))
    d2 = d1 - sig * sqrt(T-t)
    nd1 = scipy.stats.norm(0, 1).cdf(d1)
    nd2 = scipy.stats.norm(0, 1).cdf(d2)
    c = (s*nd1 - k*exp(-rf*(T-t))*nd2)
    return c

def delta(s,k,T,t,sig,rf):
    d1 = (log(s/k)+(rf+0.5*sig**2)*(T-t))/(sig*sqrt(T-t))
    nd1 = scipy.stats.norm(0, 1).cdf(d1)
    return nd1

def leverage(call,s,delta):
    lev = (s/call) * delta
    return lev

# if the time unit is year, other input should be the same as the time unit
def call_port(s_list,sig,rf):
    
    s = s_list[0]
    t = len(s)
    k_t = s
    t1 = 10/252
    t2 = 20/252
    t3 = 30/252
    t4 = 40/252
    delta_t = 0.99999/252
    delta_t1 = 0
    delta_t2 = 0
    delta_t3 = 0
    delta_t4 = 0    
    c1 = eu_call(s,k_t,t1,0,sig,rf)
    c2 = eu_call(s,k_t,t2,0,sig,rf)
    c3 = eu_call(s,k_t,t3,0,sig,rf)
    c_port = c1 + c2 + c3

    cp_list = []
    
    for i in range(t):
        delta_t1 = delta_t1 + delta_t
        delta_t2 = delta_t2 + delta_t
        delta_t3 = delta_t3 + delta_t
        delta_t4 = delta_t4 + delta_t
        
        if t1 - delta_t1 <= 0:
            delta_t1 = 10/252
            c1 = eu_call(s,k_t,t4,delta_t1,sig,rf)
            c_port = c1 + c2 + c3
            cp_list.append(c_port)
            k_t = s
            
        if t2 - delta_t2 <= 0:
            delta_t2 = 10/252
            c2 = eu_call(s,k_t,t4,delta_t2,sig,rf)
            c_port = c1 + c2 + c3
            cp_list.append(c_port)            
            k_t = s
            
        if t3 - delta_t3 <= 0:
            delta_t3 = 10/252
            c3 = eu_call(s,k_t,t4,delta_t3,sig,rf)
            c_port = c1 + c2 + c3
            cp_list.append(c_port)            
            k_t = s
        
        else:
            c1 + c2 + c3
            cp_list.append(c_port)
            
        c1 = eu_call(s,k_t,t1,delta_t1,sig,rf)
        c2 = eu_call(s,k_t,t2,delta_t2,sig,rf)
        c3 = eu_call(s,k_t,t3,delta_t3,sig,rf)
        s = s_list[i+1]
        
            


# In[127]:


21/252


# In[ ]:


reader = pd.read_csv('C:/Users/admin/Desktop/perpetual_option/coinbaseUSD.csv',header = None, iterator=True)
loop = True
chunkSize = 10000
chunks = []
while loop:
    try:
        chunk = reader.get_chunk(chunkSize)
        chunks.append(chunk)
    except StopIteration:
        loop = False
        print("Iteration is stopped.")
df = pd.concat(chunks, ignore_index=True)
#print(df)

#chunks = pd.read_csv('C:/Users/admin/Desktop/perpetual_option/coinbaseUSD.csv',header = None,iterator = True)
#chunks1 = pd.read_csv('C:/Users/admin/Desktop/perpetual_option/coinbaseUSD.csv',header = None,chunksize = 100000)
#chunk = chunks.get_chunk(5)

#db = pd.read_csv('C:/Users/admin/Desktop/perpetual_option/coinbaseUSD.csv',header = None)


# In[ ]:


df1 = df.copy()
df1[0] = df1[0].apply(f)
df_block1 = df1.iloc[0:10000,:]
df_block1.head()


# In[3]:


#a.strftime("%Y-%m-%d-%H-%M-%S")
f = lambda x: datetime.datetime.fromtimestamp(x).strftime("%Y-%m-%d")
#df1[0].apply(f)


# In[ ]:


# read data from csv file iteratively and filter it given a rule of the last trading data in each day

#df1 = chunks.get_chunk(100000)
df1 = df.copy()
df1[0] = df1[0].apply(f)
num = df1[0].unique().shape[0]
df1[df1[0]==df1[0].unique()[0]]
df_new = DataFrame()
for i in range(num):
    x = df1[df1[0]==df1[0].unique()[i]].iloc[-1,:]
    x1 = DataFrame(x).T
    df_new = pd.concat([df_new,x1],join='outer',axis=0)


# In[121]:


#df_new[x == df_new[df_new[0]==df_new[0].unique()[0]].iloc[-1,:]]
df_new[df_new[0]==df_new[0].unique()[0]]
df_new[0].unique()


# In[122]:


df_new


# In[ ]:


a = 1417412036
b = datetime.datetime.fromtimestamp(a).strftime("%Y-%m-%d")
c = datetime.datetime.fromtimestamp(a)
c + datetime.timedelta(days = 1000)


# In[ ]:





# In[ ]:





# In[68]:


df2 = df1.copy()
df2[0] = df2[0].apply(f)
df2[df2[0]=='2015-01-08'].iloc[-1,:]


# In[ ]:


df2[df2[0]=='2015-01-08'].last


# In[40]:


df[0] = df[0].apply(f)


# In[49]:


# from 2014-12-02-18-59-09 to 2015-01-08-16-26-55
ty = df.ix[5:10]
ty[0][5].split('-')


# In[17]:


time = db.iloc[7,0]
datetime.datetime.fromtimestamp(time)


# In[ ]:




