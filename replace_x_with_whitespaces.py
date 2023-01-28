import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re


path = r'E:\koodit\python\Dataprojekti\EC.xlsx'

col_list = ['Item description (local 1) (F)']

data = pd.read_excel(path, usecols=col_list)
indexi = data.index
rivit = len(indexi)
data['Tekstinsyotto'] = ''
for i in range(rivit):
	x = str(data.iloc[i,0])
	find_x = re.search(r"M?m?\d*[,.]?\d+X\d+[,.]?\d*X?x?\d*[,.]?\d*|M?m?\d*[,.]?\d+x\d+[,.]?\d*X?x?\d*[,.]?\d*",x)
	if find_x:
		print(x)
		x1 = str(find_x.group())
		x2 = x1
		x2 = x2.replace(" ","")
		x2 = x2.replace("X","x")
		x2 = x2.replace("x"," x ")
		x = x.replace(x1,x2)
		print(find_x.group())
		print("Pattern found from index",find_x.start(), find_x.end())
		data['Tekstinsyotto'][i] = x
data.to_csv(r'E:\koodit\python\Dataprojekti\EC_x.csv')