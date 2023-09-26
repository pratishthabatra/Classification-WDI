#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from openpyxl import Workbook


# In[43]:


mdata = pd.read_csv(r'C:\Users\UNOCC-data\Downloads\WDIMetadata.csv')
rdata= pd.read_csv(r'C:\Users\UNOCC-data\Downloads\RMR.csv')


# In[44]:


list(rdata.columns)
rdata.head()


# In[45]:


#nltk.download('punkt')
#nltk.download('stopwords')
def preprocess(text):
    tokens = word_tokenize(text)
    stop_words= set(stopwords.words('english'))
    tokens = [word.lower() for word in tokens if word.isalnum() and word.lower() not in stop_words]
    pptext = ' '.join(tokens)
    return pptext

def calc_cosinesim(text1, text2):
    vectorizer = TfidfVectorizer(stop_words='english')
    textvec = vectorizer.fit_transform([text1,text2])
    similarity = cosine_similarity(textvec)
    return similarity[0][1]


# In[47]:


rdata['Without pre-process Combined']=rdata.apply(lambda x:'%s_%s_%s' % (x['Risk area'],x['Description of risk area'],x['Examples of risk factors']),axis=1)
rdata['combined']=rdata.apply(lambda x:'%s_%s_%s' % (x['Risk area'],x['Description of risk area'],x['Examples of risk factors']),axis=1).apply(preprocess)


# In[49]:


rdata['Description of risk area'] = rdata['Description of risk area'].apply(preprocess)
rdata['Examples of risk factors'] = rdata['Examples of risk factors'].apply(preprocess)


# In[50]:


cut_off = -1

results1=[]
results2=[]
results3=[]
results4=[]
for _, r1row in rdata.iterrows():
    result1,result2,result3,result4 = ([],[],[],[])
    similar_texts1, similar_texts2, similar_texts3, similar_texts4 = ([],[],[],[])
    similarity_scores1, similarity_scores2,similarity_scores3,similarity_scores4 = ([],[],[],[])
    r1area = r1row['Risk area']
    r2desc = r1row['Description of risk area']
    r3exp = r1row['Examples of risk factors']
    
    for _, rrow in rdata.iterrows():
        
        r2text = rrow['Description of risk area']
        r3text = rrow['Examples of risk factors']
        risktext = rrow['Risk area']
        similarity_score1 = round(calc_cosinesim(r2desc, r2text),4)
        similarity_score2 = round(calc_cosinesim(r3exp, r3text),4)
        similarity_score3 = round(calc_cosinesim(r2desc, r3text),4)
        similarity_score4 = round(calc_cosinesim(r3exp, r2text),4)

        similarity_scores1.append(similarity_score1)
        similarity_scores2.append(similarity_score2)
        similarity_scores3.append(similarity_score3)
        similarity_scores4.append(similarity_score4)
        similar_texts1.append(risktext)
        similar_texts2.append(risktext)
        similar_texts3.append(risktext)
        similar_texts4.append(risktext)

    sorted_simdata1 = sorted(zip(similar_texts1, similarity_scores1), key=lambda x: x[1],reverse=True)
    similar_texts1, similarity_scores1= zip(*sorted_simdata1)
    sorted_simdata2 = sorted(zip(similar_texts2, similarity_scores2), key=lambda x: x[1],reverse=True)
    similar_texts2, similarity_scores2= zip(*sorted_simdata2)
    sorted_simdata3 = sorted(zip(similar_texts3, similarity_scores3), key=lambda x: x[1],reverse=True)
    similar_texts3, similarity_scores3= zip(*sorted_simdata3)
    sorted_simdata4 = sorted(zip(similar_texts4, similarity_scores4), key=lambda x: x[1],reverse=True)
    similar_texts4, similarity_scores4= zip(*sorted_simdata4)
    #result1 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(sorted_simdata1)}}         
    #results1.append(result1)
    #result2 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(sorted_simdata2)}}        
    #results2.append(result2)
    #result3 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(sorted_simdata3)}}             
    #results3.append(result3)
    #result4 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(sorted_simdata4)}}             
    #results4.append(result4)
    result1 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(similar_texts1)},**{f'Associated Risks areas score {i+1}': score for i,score in enumerate(similarity_scores1)}}
    results1.append(result1)
    result2 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(similar_texts2)},**{f'Associated Risks areas score {i+1}': score for i,score in enumerate(similarity_scores2)}}
    results2.append(result2)
    result3 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(similar_texts3)},**{f'Associated Risks areas score {i+1}': score for i,score in enumerate(similarity_scores3)}}
    results3.append(result3)
    result4 = {'Risk area':r1area,**{f'Similar Risks areas {i+1}': text for i,text in enumerate(similar_texts4)},**{f'Associated Risks areas score {i+1}': score for i,score in enumerate(similarity_scores4)}}
    results4.append(result4)

results1_df = pd.DataFrame(results1)
results2_df = pd.DataFrame(results2)
results3_df = pd.DataFrame(results3)
results4_df = pd.DataFrame(results4)
results1_df.head()


# In[51]:


writer = pd.ExcelWriter('RMR-within_similarity-sep.xlsx', engine='openpyxl')


# In[52]:


results1_df.to_excel(writer, sheet_name='Desc-Desc', index=False)
results2_df.to_excel(writer, sheet_name='Examples-Examples', index=False)
results3_df.to_excel(writer, sheet_name='Desc-Examples', index=False)
results4_df.to_excel(writer, sheet_name='Examples-Desc', index=False)


# In[54]:


writer.book.save('RMR-within_similarity-sep.xlsx')


# In[ ]:




