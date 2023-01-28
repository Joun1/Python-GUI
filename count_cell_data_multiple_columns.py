import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re

path = r'E:\koodit\python\Dataprojekti\first list item descriptions.xlsx'

col_list = ['Item Group','Product sub Category','Product Type Name','MrtlSubGroup','Cert Purchase Doc Description','Item type 1']

data = pd.read_excel(path, usecols=col_list)

new_df = pd.DataFrame()

for i in range(len(col_list)):
    text = list(data[col_list[i]].values.astype(str))

    text_len = len(text)
#    data[col_list[i]] = ''

    for j in range(text_len):
        data[col_list[i]][j] = text[j]
    print(text)

    df1 = pd.DataFrame(text,columns = [col_list[i]])
    df1['count'] = 1
    df1 = df1.groupby([col_list[i]])['count'].count()
    df1 = df1.reset_index()
    
    columnit = df1[[col_list[i],'count']]
    new_df = new_df.append(columnit,ignore_index = False)
    new_df.rename(columns={'count': 'count'+str(i)},inplace=True)
    print(new_df)
new_df = new_df.apply(lambda x: pd.Series(x.dropna().values))
new_df.to_excel(r'E:\koodit\python\Dataprojekti\first list item descriptions_words_modified.xlsx',index=False)