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
from sentence_transformers import SentenceTransformer
from tqdm import tqdm
import torch
from openpyxl import Workbook


# In[2]:


mdata = pd.read_csv(r'C:\Users\UNOCC-data\Downloads\WDIMetadata.csv')
rdata= pd.read_csv(r'C:\Users\UNOCC-data\Downloads\RMR.csv')


# In[3]:


rdata.head()


# In[4]:


rdata['RiskNum'] = range(1,len(rdata['Risk area'])+1,1)
rdata.head()


# In[5]:


mdata.dropna(subset=['Indicator Name'], inplace=True)
mdata.shape


# In[6]:


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


# In[7]:


mdata['combined']=mdata.apply(lambda x:'%s_%s_%s_%s' % (x['Indicator Name'],x['Short definition'],x['Long definition'],x['Topic']),axis=1).apply(preprocess)
mdata['ppshortdef'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Short definition']),axis=1).apply(preprocess)
mdata['pplongdef'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Long definition']),axis=1).apply(preprocess)
mdata['pptopic'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Topic']),axis=1).apply(preprocess)
mdata['pplongdeftopic'] = mdata.apply(lambda x:'%s_%s_%s' % (x['Indicator Name'],x['Long definition'],x['Topic']),axis=1).apply(preprocess)
mdata['ppshortdeftopic'] = mdata.apply(lambda x:'%s_%s_%s' % (x['Indicator Name'],x['Short definition'],x['Topic']),axis=1).apply(preprocess)


# In[8]:


rdata['combined']=rdata.apply(lambda x:'%s_%s_%s' % (x['Risk area'],x['Description of risk area'],x['Examples of risk factors']),axis=1).apply(preprocess)
rdata['ppDesc']=rdata.apply(lambda x:'%s_%s' % (x['Risk area'],x['Description of risk area']),axis=1).apply(preprocess)
rdata['ppExamp']=rdata.apply(lambda x:'%s_%s' % (x['Risk area'],x['Examples of risk factors']),axis=1).apply(preprocess)


# In[92]:


cut_off = 0.005
n = 5 # top 5 risk areas

results1=[]
results2=[]
for mindex, mrow in mdata.iterrows():
    result1 = []
    result2 = []
    similar_texts = []
    similarity_scores = []
    similar_risknum = []
    mcode = mrow['Code']
    mtopic = mrow['Topic']
    #meta_text = mrow['combined']
    #meta_text = mrow['ppshortdef']
    #meta_text = mrow['pplongdef']
    #meta_text = mrow['pptopic']
    #meta_text = mrow['ppshortdeftopic']
    meta_text = mrow['pplongdeftopic']
    itext = mrow['Indicator Name']
    
    for rindex, rrow in rdata.iterrows():
        
        #rtext = rrow['combined']
        rtext = rrow['ppDesc']
        #rtext = rrow['ppExamp']
        risktext = rrow['Risk area']
        risktextnum = rrow['RiskNum']
        similarity_score = round(calc_cosinesim(meta_text, rtext),4)
        #similarity_score2 = calc_cosinesim(mtopic, rtext)
        
        if similarity_score >= cut_off:
            similarity_scores.append(similarity_score)
            similar_texts.append(risktext)
            similar_risknum.append(risktextnum)

    if similarity_scores:
        sorted_simdata = sorted(zip(similar_texts, similarity_scores,similar_risknum), key=lambda x: x[1],reverse=True)
        similar_texts, similarity_scores, similar_risknum= zip(*sorted_simdata)
    if len(similarity_scores) > n:
        similar_texts = similar_texts[0:n]
        similarity_scores = similarity_scores[0:n]
        similar_risknum = similar_risknum[0:n]
    result1 = {'Code':mcode,'Indicator': itext,**{f'Similar Risk areas {i+1}': text for i,text in enumerate(similar_texts)},**{f'Associated Risks areas score {i+1}': score for i,score in enumerate(similarity_scores)}}         
    result2 = {'Code':mcode,'Indicator': itext,**{f'Similar Risk Num {i+1}': np.round(j) for i,j in enumerate(similar_risknum)},**{f'Associated Risks areas score {i+1}': score for i,score in enumerate(similarity_scores)}}         

    results1.append(result1)
    results2.append(result2)

results1_df = pd.DataFrame(results1)
results2_df = pd.DataFrame(results2)
results1_df.head()


# In[94]:


results1_df['Similar Risk areas 1'].isnull().sum()


# In[95]:


results2_df.head()


# In[96]:


writer = pd.ExcelWriter('WDI-RMR-LTD.xlsx', engine='openpyxl')


# In[97]:


results1_df.to_excel(writer, sheet_name='Main', index=False)
results2_df.to_excel(writer, sheet_name='Numbers', index=False)


# In[98]:


writer.book.save('WDI-RMR-LTD.xlsx')
writer = pd.ExcelWriter('WDI-RMR-LTD.xlsx', engine='openpyxl', mode='a')


# In[10]:


import os
os.getcwd()
path = "C://Users//UNOCC-data//Downloads//WDIRMR"
dir_list = os.listdir(path)
 
print("Files and directories in '", path, "' :")
print(dir_list)


# In[93]:


os.getcwd()


# In[60]:


os.chdir('C:\\Users\\UNOCC-data\\Downloads\\WDIRMR')


# In[171]:


import glob

path = r'C:\Users\UNOCC-data\Downloads\WDIRMR'
all_files = glob.glob(os.path.join(path, "*.xlsx"))
li = []
sheetname="Numbers"
for filename in all_files:
    df = pd.read_excel(filename, header=0,sheet_name=sheetname)
    li.append({'filename': filename, 'data': df})


# In[213]:


#li[0]['data']['Penalty scores'] =penalty_totals
li[1]['data']


# In[227]:


li[0]['data'].iloc[:,2:7]


# In[376]:


from statistics import mean 
current_rows = li[0]['data'].iloc[:, 2:7].values.tolist() 

penalty_avg = [] 
step1_avg = [] 
step2_avg = []
penalty_avg.append(0)
step1_avg.append(0)
step2_avg.append(0)
blanks = []
exact_match = []
exact_match.append(0)
blanks.append(li[0]['data'].iloc[:,2].isnull().sum())
allpenalty_counts=[]

for j in range(1,18):
    step1_totals=[]
    step2_totals=[]
    penalty_totals=[]
    li[j]['data']['Penalty Scores']=[None] * len(li[0]['data'].iloc[:,2])
    li[j]['data']['Different risks']=[None] * len(li[0]['data'].iloc[:,2])
    li[j]['data']['Unordered risks']=[None] * len(li[0]['data'].iloc[:,2])
    
    for i in range(len(li[0]['data'].iloc[:,2])):
        step1_move, step2_move, penalty_score=(0,0,0)
        current_row = []
        desired_row = []
        desired_rows = []
        current_row = current_rows[i]
        desired_rows = li[j]['data'].iloc[:, 2:7].values.tolist()
        desired_row =desired_rows[i]
        step1_move, step2_move, penalty_score = calculate_moves(current_row, desired_row)
        #li[j]['data']['Penalty Scores'].append(step1_move)
        #li[j]['data']['Different risks'].append(step2_move)
        #li[j]['data']['Unordered risks'].append(penalty_score)
        #li[j]['data']['Penalty Scores'][i]=penalty_score
        #li[j]['data']['Different risks'][i]=step1_move
        #li[j]['data']['Unordered risks'][i]=step2_move
        penalty_totals.append(penalty_score)
        step1_totals.append(step1_move)
        step2_totals.append(step2_move)
        
    li[j]['data']['Penalty Scores']=penalty_totals
    li[j]['data']['Different risks']=step1_totals
    li[j]['data']['Unordered risks']=step2_totals
    
    penalty_avg.append(mean(li[j]['data']['Penalty Scores']))
    step1_avg.append(mean(li[j]['data']['Different risks']))
    step2_avg.append(mean(li[j]['data']['Unordered risks']))
    blanks.append(li[j]['data'].iloc[:,2].isnull().sum())
    exact_match.append((li[j]['data']['Penalty Scores'][:]==0).sum())
    allpenalty_counts.append(li[j]['data'].groupby(by='Penalty Scores').size())




# In[377]:


row_names=['Combined','Short definition', 'Long definition','Topic', 'Topic + Short definition','Topic + Long definition']
column_names=['Combined', 'Description of risk areas','Examples of risk areas']
penaltymatrix= pd.DataFrame([penalty_avg[i:i+3] for i in range(0,len(penalty_avg),3)],columns=column_names, index=row_names)
differentrisksmatrix= pd.DataFrame([step1_avg[i:i+3] for i in range(0,len(step1_avg),3)],columns=column_names, index=row_names)
differentorderedrisksmatrix= pd.DataFrame([step2_avg[i:i+3] for i in range(0,len(step2_avg),3)],columns=column_names, index=row_names)
exactmatchmatrix= pd.DataFrame([exact_match[i:i+3] for i in range(0,len(exact_match),3)],columns=column_names, index=row_names)
blanksmatrix= pd.DataFrame([blanks[i:i+3] for i in range(0,len(blanks),3)],columns=column_names, index=row_names) 
allpenalty_countsmatrix= pd.DataFrame(allpenalty_counts,index=['CD','CE','SC','SD','SE','LC','LD','LE','TC','TD','TE','STC','STD','STE','LTC','LTD','LTE'])


# In[378]:


dd=li[1]['data'].groupby(by='Penalty Scores').size()
print(dd)
#len(dd)
#(li[1]['data']['Penalty Scores'][:]==0).sum()
#li[1]['data']['Penalty Scores']


# In[379]:


#li[1]['data'].iloc[:,12].count("2")
#allpenalty_counts= li[1]['data'].groupby(by='Penalty Scores').size()
exact_match
#penalty_avg
#penaltymatrix
#differentrisksmatrix
#allpenalty_counts


# In[380]:


writer = pd.ExcelWriter('WDI-RMR-sensitivityanalysis.xlsx', engine='openpyxl')


# In[381]:


penaltymatrix.to_excel(writer, sheet_name='Penalty', index=True)
differentrisksmatrix.to_excel(writer, sheet_name='Different risks penalty', index=True)
differentorderedrisksmatrix.to_excel(writer, sheet_name='Unordered risks penalty', index=True)
exactmatchmatrix.to_excel(writer, sheet_name='Exact matches', index=True)
allpenalty_countsmatrix.to_excel(writer, sheet_name='All_penalties', index=True)
blanksmatrix.to_excel(writer, sheet_name='Unassigned', index=True)


# In[382]:


writer.book.save('WDI-RMR-sensitivityanalysis.xlsx')
writer = pd.ExcelWriter('WDI-RMR-sensitivityanalysis.xlsx', engine='openpyxl', mode='a')

