#from example import *
import pandas as pd
import PySimpleGUI as sg
import re
from example import event,values

def display_excel_file(excel_file_path):
	df = pd.read_excel(excel_file_path)
	filename = excel_file_path
	print(values["-IN-"])
	x = re.search(r"/\w*\.",filename)
	print(x,"=X")
	outputfilename = str(x.group())
	outputfilename = outputfilename[:-1] + "_modified.csv"
	print(outputfilename, "outputfilename")
	sg.popup_scrolled(df.index,df,title=filename)