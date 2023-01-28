import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re

path = r'E:\koodit\python\Dataprojekti\EC.xlsx'

col_list = ['Item description (local 1)']

data = pd.read_excel(path, usecols=col_list)

new_df = pd.DataFrame()

for i in range(len(col_list)):
    text = " ".join(review for review in data[col_list[i]].astype(str))
    text = text.split()

    for j in text:
        j = j.strip()
        if len(j) == 0:
            text.pop(text.index(j))

    text_len = len(text)
    data[col_list[i]] = ''

    for j in range(text_len):
        data[col_list[i]][j] = text[j]

    df1 = pd.DataFrame(text,columns = [col_list[i]])
    df1['count'] = 1
    df1 = df1.groupby([col_list[i]])['count'].count()
    df1 = df1.reset_index()
    
    columnit = df1[[col_list[i],'count']]
    new_df = new_df.append(columnit,ignore_index = False)
    new_df.rename(columns={'count': 'count'+str(i)},inplace=True)

new_df = new_df.apply(lambda x: pd.Series(x.dropna().values))
new_df.to_excel(r'E:\koodit\python\Dataprojekti\EC_modified.xlsx',index=False)