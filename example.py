import pandas as pd
import PySimpleGUI as sg
import re

pd.set_option('display.max_columns', None)
#pd.set_option('display.width', 320)

def count_column_terms(path,columnit):
	col_list = [columnit]
	data = pd.read_excel(path, usecols=col_list)
	new_df = pd.DataFrame()

	for i in range(len(col_list)):
		text = list(data[col_list[i]].values.astype(str))

	text_len = len(text)
#    data[col_list[i]] = ''

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

	outputfilename = output_filename(path)
	filunimi = str(values["-OUT-"])+"/exports"+str(outputfilename)
	filunimi = filunimi.replace("_modified.csv","_count_column_terms_modified.csv")
	new_df.to_csv(filunimi,index=False)

def count_cell_terms(path,columnit):
	col_list = [columnit]
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
    
		columni = df1[[col_list[i],'count']]
		new_df = new_df.append(columni,ignore_index = False)
		new_df.rename(columns={'count': 'count'+str(i)},inplace=True)

	outputfilename = output_filename(path)
	filunimi = str(values["-OUT-"])+"/exports"+str(outputfilename)
	new_df = new_df.apply(lambda x: pd.Series(x.dropna().values))
	filunimi = filunimi.replace("_modified.csv","_count_cell_terms_modified.csv")
	new_df.to_csv(filunimi,index=False)

def output_filename(excel_file_path):
	print(excel_file_path)
	x = re.search(r"/\w*\.",excel_file_path)
	outputfilename = str(x.group())
	outputfilename = outputfilename[:-1] + "_modified.csv"
	return outputfilename

def separate_x(path,columnit):

	col_list = [columnit]
	data = pd.read_excel(path, usecols=col_list)
	indexi = data.index

	outputfilename = output_filename(path)
	rivit = len(indexi)
	data['Tekstinsyotto'] = ''
	for i in range(rivit):
		x = str(data.iloc[i,0])
		find_x = re.search(r"M?m?\d*[,.]?\d+X\d+[,.]?\d*X?x?\d*[,.]?\d*|M?m?\d*[,.]?\d+x\d+[,.]?\d*X?x?\d*[,.]?\d*",x)
		if find_x:
			x1 = str(find_x.group())
			x2 = x1
			x2 = x2.replace(" ","")
			x2 = x2.replace("X","x")
			x2 = x2.replace("x"," x ")
			x = x.replace(x1,x2)
			data['Tekstinsyotto'][i] = x
	filunimi = str(values["-OUT-"])+"/exports"+str(outputfilename)
	filunimi = filunimi.replace("_modified.csv","_separate_x_modified.csv")
	data.to_csv(filunimi)

def is_valid_path(filepath):
	if filepath:
		return True
	sg.popup_error("Filepath not correct")
	return False

def display_excel_file(excel_file_path):
	df = pd.read_excel(excel_file_path)
	filename = excel_file_path
	x = re.search(r"/\w*\.",filename)
	outputfilename = str(x.group())
	outputfilename = outputfilename[:-1] + "_modified.csv"
	sg.popup_scrolled(df.columns.values,title=filename)

layout = [
		[sg.Text("Input File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),))],
		[sg.Text("Output Folder:"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
		[sg.Button("Display available columns from chosen")],
		[sg.Button("Separate \"X\""),sg.Text("Select column:"),sg.Input(key="-IN-")],
		[sg.Button("Count terms from column"),sg.Text("Select column:"),sg.Input(key="-IN-")],
		[sg.Button("Count column terms"),sg.Text("Select column:"),sg.Input(key="-IN-")],
		[sg.Exit()],
],

window = sg.Window("Display Excel File", layout)

while True:
	event, values = window.read()
	print(event,values)
	if event in (sg.WINDOW_CLOSED, "Exit"):
		break
	if event == "Display available columns from chosen":
		if is_valid_path(values["-IN-"]):
			display_excel_file(values["-IN-"])
	if event == "Separate \"X\"":
		separate_x(str(values["-IN-"]),str(values["-IN-1"]))
		print(values["-IN-"],values["-IN-1"])
	if event == "Count terms from column":
		count_cell_terms(str(values["-IN-"]),str(values["-IN-2"]))
		print(values["-IN-"],values["-IN-2"])
	if event == "Count column terms":
		count_column_terms(str(values["-IN-"]),str(values["-IN-3"]))
		print(values["-IN-"],values["-IN-3"])
window.close()