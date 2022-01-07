import zipfile
import csv
from bs4 import BeautifulSoup as Soup
import tkinter as tk
from tkinter import filedialog
import re

# to show file selection dialog box
root = tk.Tk()
root.withdraw()
paths = filedialog.askopenfilenames()
root.update()

with open('/'.join(paths[0].split('/')[0:-1])+'/output.csv', 'w', newline='', encoding='utf-8-sig') as f:
	csvw = csv.writer(f)
	
	# loop through each selected file
	for path in paths:
		
		# Writing header line with the selected file's filename
		csvw.writerow([path.split('/')[-1], ''])

		# .docx files unzipping
		unzip = zipfile.ZipFile(path)
		comments = Soup(unzip.read('word/comments.xml'), 'lxml')
		
		# Selected document file preprocessing to handle multi-paragraph and nested comments and unzipping it into a string first
		doc = unzip.read('word/document.xml').decode()

		# Finding all the comment start and end locations and storing them in dictionaries keyed on the unique ID for each comment
		start_loc = {x.group(1): x.start() for x in re.finditer(r'<w:commentRangeStart.*?w:id="(.*?)"', doc)}
		end_loc = {x.group(1): x.end() for x in re.finditer(r'<w:commentRangeEnd.*?w:id="(.*?)".*?>', doc)}
		
		# looping through all the comments in the comments.xml file
		for c in comments.find_all('w:comment'):
			c_id = c.attrs['w:id']

			# Using the previous found locations to extract the xml fragment from the document for each comment ID, adding spaces to separate any paragraphs in multi-paragraph comments
			xml = re.sub(r'(<w:p .*?>)', r'\1 ', doc[start_loc[c_id]:end_loc[c_id] + 1])
			
			# Parsing the XML fragment and writing it to file along with the label text
			csvw.writerow([''.join(c.findAll(text=True)), ''.join(Soup(xml, 'lxml').findAll(text=True))])

	unzip.close()