#!/usr/bin/env python
# coding: utf-8

# In[4]:


pip install -U sentence-transformers


# In[2]:


pip install sentence-transformers


# In[5]:


import os
os. getcwd()


# In[6]:


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


# In[18]:


mdata = pd.read_csv(r'/Users/pratishthabatra/Downloads/WDIMetadata.csv')
rdata= pd.read_csv(r'/Users/pratishthabatra/Downloads/RMR.csv')


# In[19]:


rdata['RiskNum'] = range(1,len(rdata['Risk area'])+1,1)
rdata.head()


# In[20]:


from transformers import BertTokenizer, BertModel

bertmodel = "bert-base-uncased"
tokenizer = BertTokenizer.from_pretrained(bertmodel)
model = BertModel.from_pretrained(bertmodel)


# In[21]:


def encode_text(text):
    tokens  = tokenizer(text, return_tensors="pt", padding=True, truncation=True)
    with torch.no_grad():
        outputs = model(**tokens)
    return outputs.last_hidden_state.mean(dim=1).squeeze().numpy()


# In[22]:


mdata['combined'] = mdata.apply(lambda x:'%s_%s_%s_%s' % (x['Indicator Name'],x['Short definition'],x['Long definition'],x['Topic']),axis=1).apply(encode_text)
mdata['ppshortdef'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Short definition']),axis=1).apply(encode_text)
mdata['pplongdef'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Long definition']),axis=1).apply(encode_text)
mdata['pptopic'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Topic']),axis=1).apply(encode_text)
mdata['pplongdeftopic'] = mdata.apply(lambda x:'%s_%s_%s' % (x['Indicator Name'],x['Long definition'],x['Topic']),axis=1).apply(encode_text)
mdata['ppshortdeftopic'] = mdata.apply(lambda x:'%s_%s_%s' % (x['Indicator Name'],x['Short definition'],x['Topic']),axis=1).apply(encode_text)


# In[23]:


rdata['combined']=rdata.apply(lambda x:'%s_%s_%s' % (x['Risk area'],x['Description of risk area'],x['Examples of risk factors']),axis=1).apply(encode_text)
rdata['ppDesc']=rdata.apply(lambda x:'%s_%s' % (x['Risk area'],x['Description of risk area']),axis=1).apply(encode_text)
rdata['ppExamp']=rdata.apply(lambda x:'%s_%s' % (x['Risk area'],x['Examples of risk factors']),axis=1).apply(encode_text)


# In[24]:


def calc_cosinesim(emb1, emb2):
    similarity = cosine_similarity(emb1,emb2)[0][0]
    return similarity


# In[120]:


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
    meta_text = mrow['combined'].reshape(1,-1)
    #meta_text = mrow['ppshortdef'].reshape(1,-1)
    #meta_text = mrow['pplongdef'].reshape(1,-1)
    #meta_text = mrow['pptopic'].reshape(1,-1)
    #meta_text = mrow['ppshortdeftopic'].reshape(1,-1)
    #meta_text = mrow['pplongdeftopic'].reshape(1,-1)
    itext = mrow['Indicator Name']
    
    for rindex, rrow in rdata.iterrows():
        
        rtext = rrow['combined'].reshape(1,-1)
        #rtext = rrow['ppDesc'].reshape(1,-1)
        #rtext = rrow['ppExamp'].reshape(1,-1)
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


# In[121]:


results1_df['Similar Risk areas 1'].isnull().sum()


# In[122]:


results2_df.head()


# In[123]:


writer = pd.ExcelWriter('BERTWDI-RMR-CC.xlsx', engine='openpyxl')


# In[124]:


results1_df.to_excel(writer, sheet_name='Main', index=False)
results2_df.to_excel(writer, sheet_name='Numbers', index=False)

writer.book.save('BERTWDI-RMR-CC.xlsx')
writer = pd.ExcelWriter('BERTWDI-RMR-CC.xlsx', engine='openpyxl', mode='a')


# In[126]:


os.getcwd()


# In[129]:


os.getcwd()
path = "/Users/pratishthabatra/Downloads/BERTWDIRMR"
dir_list = os.listdir(path)
 
print("Files and directories in '", path, "' :")
print(dir_list)


# In[130]:


os.getcwd()


# In[131]:


os.chdir('/Users/pratishthabatra/Downloads/BERTWDIRMR')


# In[158]:


import glob

path = r'/Users/pratishthabatra/Downloads/BERTWDIRMR'
#files = sorted(glob.glob('/Users/pratishthabatra/Downloads/BERTWDIRMR*.xlsx'))
all_files = sorted(glob.glob(os.path.join(path, "*.xlsx")))
li = []
sheetname="Numbers"
for filename in all_files:
    df = pd.read_excel(filename, header=0,sheet_name=sheetname)
    li.append({'filename': filename, 'data': df})



# In[161]:


li[2]['filename']


# In[162]:


li[0]['filename']


# In[164]:


# Function to calculate the number of unordered or different risk areas
def calculate_moves(row1, row2):
    # Create dictionaries to store the positions of risk areas in row2
    dict2 = {val: index for index, val in enumerate(row2)}
    penalty = 0
    # Initialize variables to count moves for each step
    step1_moves = 0  # Element Removal
    step2_moves = 0  # Element Swap
    current_copy = row1
    # Step 1: Element Removal
    for element in current_copy:
        if element not in dict2:
            step1_moves += 1  # Count as a move to remove the element

    # Step 2: Element Swap
    if  current_copy:
        for i in range(min(len(current_copy), len(desired_rows))):
            element = current_copy[i]
            if element in dict2:
                index2 = dict2[element]
        # If the element is not in its desired position, swap it (Step 2)
                if i != index2:
                    current_copy[i], current_copy[index2] = current_copy[index2], current_copy[i]
                    step2_moves += 1  # Count a move to swap the element
            
    if step1_moves==0 and step2_moves==0:
        penalty=0
    elif step1_moves == 0 and step2_moves > 0:
        penalty=1
    elif step1_moves >0 and step2_moves >0:
        penalty = 2
    elif step1_moves > 0 and step2_moves == 0:
        penalty =3
  
    return step1_moves, step2_moves, penalty


# In[165]:


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



# In[166]:


row_names=['Combined','Short definition', 'Long definition','Topic', 'Topic + Short definition','Topic + Long definition']
column_names=['Combined', 'Description of risk areas','Examples of risk areas']
penaltymatrix= pd.DataFrame([penalty_avg[i:i+3] for i in range(0,len(penalty_avg),3)],columns=column_names, index=row_names)
differentrisksmatrix= pd.DataFrame([step1_avg[i:i+3] for i in range(0,len(step1_avg),3)],columns=column_names, index=row_names)
differentorderedrisksmatrix= pd.DataFrame([step2_avg[i:i+3] for i in range(0,len(step2_avg),3)],columns=column_names, index=row_names)
exactmatchmatrix= pd.DataFrame([exact_match[i:i+3] for i in range(0,len(exact_match),3)],columns=column_names, index=row_names)
blanksmatrix= pd.DataFrame([blanks[i:i+3] for i in range(0,len(blanks),3)],columns=column_names, index=row_names) 
allpenalty_countsmatrix= pd.DataFrame(allpenalty_counts,index=['CD','CE','SC','SD','SE','LC','LD','LE','TC','TD','TE','STC','STD','STE','LTC','LTD','LTE'])


# In[167]:


writer = pd.ExcelWriter('BERTWDI-RMR-sensitivityanalysis.xlsx', engine='openpyxl')

penaltymatrix.to_excel(writer, sheet_name='Penalty', index=True)
differentrisksmatrix.to_excel(writer, sheet_name='Different risks penalty', index=True)
differentorderedrisksmatrix.to_excel(writer, sheet_name='Unordered risks penalty', index=True)
exactmatchmatrix.to_excel(writer, sheet_name='Exact matches', index=True)
allpenalty_countsmatrix.to_excel(writer, sheet_name='All_penalties', index=True)
blanksmatrix.to_excel(writer, sheet_name='Unassigned', index=True)

writer.book.save('BERTWDI-RMR-sensitivityanalysis.xlsx')
writer = pd.ExcelWriter('BERTWDI-RMR-sensitivityanalysis.xlsx', engine='openpyxl', mode='a')



# In[ ]:




