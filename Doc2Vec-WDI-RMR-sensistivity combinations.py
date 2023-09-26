#!/usr/bin/env python
# coding: utf-8

# In[5]:


from nltk.tokenize import sent_tokenize, word_tokenize
import gensim
from gensim.models import Word2Vec
from gensim.models.doc2vec import Doc2Vec, TaggedDocument


# In[6]:


import pandas as pd
import numpy as np
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# In[46]:


mdata = pd.read_csv(r'C:\Users\UNOCC-data\Downloads\WDIMetadata.csv')
rdata= pd.read_csv(r'C:\Users\UNOCC-data\Downloads\RMR.csv')
mdata.head()


# In[47]:


rdata['RiskNum'] = range(1,len(rdata['Risk area'])+1,1)
rdata.head()


# In[48]:


def preprocess(text):
    tokens =text.split()
    stop_words= set(stopwords.words('english'))
    tokens = [word.lower() for word in tokens if word.isalnum() and word.lower() not in stop_words]
    pptext = ' '.join(tokens)
    return pptext


# In[49]:


mdata['combined']=mdata.apply(lambda x:'%s_%s_%s_%s' % (x['Indicator Name'],x['Short definition'],x['Long definition'],x['Topic']),axis=1).apply(preprocess)
mdata['ppshortdef'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Short definition']),axis=1).apply(preprocess)
mdata['pplongdef'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Long definition']),axis=1).apply(preprocess)
mdata['pptopic'] = mdata.apply(lambda x:'%s_%s' % (x['Indicator Name'],x['Topic']),axis=1).apply(preprocess)
mdata['pplongdeftopic'] = mdata.apply(lambda x:'%s_%s_%s' % (x['Indicator Name'],x['Long definition'],x['Topic']),axis=1).apply(preprocess)
mdata['ppshortdeftopic'] = mdata.apply(lambda x:'%s_%s_%s' % (x['Indicator Name'],x['Short definition'],x['Topic']),axis=1).apply(preprocess)


# In[50]:


rdata['combined']=rdata.apply(lambda x:'%s_%s_%s' % (x['Risk area'],x['Description of risk area'],x['Examples of risk factors']),axis=1).apply(preprocess)
rdata['ppDesc']=rdata.apply(lambda x:'%s_%s' % (x['Risk area'],x['Description of risk area']),axis=1).apply(preprocess)
rdata['ppExamp']=rdata.apply(lambda x:'%s_%s' % (x['Risk area'],x['Examples of risk factors']),axis=1).apply(preprocess)


# In[130]:


# Preparing the data for Doc2Vec training
tagged_data = [TaggedDocument(words=text.split(), tags=[str(i)]) for i, text in enumerate(rdata['combined'])]

model.random.seed(123)
# Training the Doc2Vec model : Do this once and save the model using model.save()
#model = Doc2Vec(vector_size=200, window=2, min_count=1, workers=4, epochs=20)
#model.build_vocab(tagged_data)
#model.train(tagged_data, total_examples=model.corpus_count, epochs=model.epochs)

# Loading the saved model
model = Doc2Vec.load("RMR-similar_sentence.model")

cut_off = 0.05
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
    meta_text = mrow['combined']
    #meta_text = mrow['ppshortdef']
    #meta_text = mrow['pplongdef']
    #meta_text = mrow['pptopic']
    #meta_text = mrow['ppshortdeftopic']
    #meta_text = mrow['pplongdeftopic']
    itext = mrow['Indicator Name']
    model.random.seed(123)
    meta_vector = model.infer_vector(meta_text.split())

    
    for rindex, rrow in rdata.iterrows():
        
        rtext = rrow['combined']
        #rtext = rrow['ppDesc']
        #rtext = rrow['ppExamp']
        risktext = rrow['Risk area']
        risktextnum = rrow['RiskNum']
        model.random.seed(123)
        r_vector = model.infer_vector(rtext.split())
        similarity_score = cosine_similarity([meta_vector], [r_vector])[0][0]
        similarity_score = round(similarity_score,4)
        
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


# In[108]:


results1_df['Similar Risk areas 1'].isnull().sum()


# In[62]:


import os
os.getcwd()


# In[63]:


os.chdir('C:\\Users\\UNOCC-data\\Downloads\\WDIRMR-Doc2Vec')


# In[68]:


# Right after training, save this model
#model.save("RMR-similar_sentence.model")
# Load the saved model
#model = Doc2Vec.load("RMR-similar_sentence.model")


# In[131]:


writer = pd.ExcelWriter('WDI-RMR-CC-Doc2vec.xlsx', engine='openpyxl')


# In[132]:


results1_df.to_excel(writer, sheet_name='Main', index=False)
results2_df.to_excel(writer, sheet_name='Numbers', index=False)
writer.book.save('WDI-RMR-CC-Doc2vec.xlsx')
writer = pd.ExcelWriter('WDI-RMR-CC-Doc2vec.xlsx', engine='openpyxl', mode='a')


# In[134]:


import os
os.getcwd()
path = "C://Users//UNOCC-data//Downloads//WDIRMR-Doc2Vec"
dir_list = os.listdir(path)
 
print("Files and directories in '", path, "' :")
print(dir_list)


# In[135]:


import glob

path = r'C:\Users\UNOCC-data\Downloads\WDIRMR-Doc2Vec'
all_files = sorted(glob.glob(os.path.join(path, "*.xlsx")))
li = []
sheetname="Numbers"
for filename in all_files:
    df = pd.read_excel(filename, header=0,sheet_name=sheetname)
    li.append({'filename': filename, 'data': df})


# In[136]:


#li[0]['data']['Penalty scores'] =penalty_totals
li[1]['data']


# In[137]:


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


# In[138]:


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




# In[139]:


row_names=['Combined','Short definition', 'Long definition','Topic', 'Topic + Short definition','Topic + Long definition']
column_names=['Combined', 'Description of risk areas','Examples of risk areas']
penaltymatrix= pd.DataFrame([penalty_avg[i:i+3] for i in range(0,len(penalty_avg),3)],columns=column_names, index=row_names)
differentrisksmatrix= pd.DataFrame([step1_avg[i:i+3] for i in range(0,len(step1_avg),3)],columns=column_names, index=row_names)
differentorderedrisksmatrix= pd.DataFrame([step2_avg[i:i+3] for i in range(0,len(step2_avg),3)],columns=column_names, index=row_names)
exactmatchmatrix= pd.DataFrame([exact_match[i:i+3] for i in range(0,len(exact_match),3)],columns=column_names, index=row_names)
blanksmatrix= pd.DataFrame([blanks[i:i+3] for i in range(0,len(blanks),3)],columns=column_names, index=row_names) 
allpenalty_countsmatrix= pd.DataFrame(allpenalty_counts,index=['CD','CE','SC','SD','SE','LC','LD','LE','TC','TD','TE','STC','STD','STE','LTC','LTD','LTE'])


# In[140]:


writer = pd.ExcelWriter('WDI-RMR-Doc2Vec-sensitivityanalysis.xlsx', engine='openpyxl')


# In[141]:


penaltymatrix.to_excel(writer, sheet_name='Penalty', index=True)
differentrisksmatrix.to_excel(writer, sheet_name='Different risks penalty', index=True)
differentorderedrisksmatrix.to_excel(writer, sheet_name='Unordered risks penalty', index=True)
exactmatchmatrix.to_excel(writer, sheet_name='Exact matches', index=True)
allpenalty_countsmatrix.to_excel(writer, sheet_name='All_penalties', index=True)
blanksmatrix.to_excel(writer, sheet_name='Unassigned', index=True)


# In[142]:


writer.book.save('WDI-RMR-Doc2Vec-sensitivityanalysis.xlsx')
writer = pd.ExcelWriter('WDI-RMR-Doc2Vec-sensitivityanalysis.xlsx', engine='openpyxl', mode='a')


# In[143]:


dd=li[1]['data'].groupby(by='Penalty Scores').size()
print(dd)


# In[147]:


li[1]


# In[ ]:




