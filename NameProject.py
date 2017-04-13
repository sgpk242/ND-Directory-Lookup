import xlrd
import datetime, time, threading, textwrap, re, os
from bs4 import BeautifulSoup
from urllib2 import urlopen
from xlrd.sheet import ctype_text
import xlwt

def findHTML(name):
	html = urlopen('https://apps.nd.edu/webdirectory/directory.cfm?specificity=contains&cn={}&Submit=Submit'.format(name)).read()
        soup = BeautifulSoup(html, 'lxml')
        div = soup.find('div', align="center")
        table = div.table
        section = table.find('td', colspan="2").table
        s = section.findAll('a')
	return s

workbook = xlrd.open_workbook('NameProject.xlsx')
new_wb = xlwt.Workbook()
tot_sheets = len(workbook.sheet_names())

for sheet_num in range(0, tot_sheets):
	sheet = workbook.sheet_by_index(sheet_num)
	ws = new_wb.add_sheet(workbook.sheet_names()[sheet_num])
	ws.write(0, 0, workbook.sheet_names()[sheet_num] + ' - NOT STARTED TRAVEL REGISTRATION')

	for row_idx in range(2, sheet.nrows):
		name = []
		for col_idx in range(0, sheet.ncols):
			cell_obj = sheet.cell(row_idx, col_idx)
			name.append('{}'.format(cell_obj.value))
		name_str = '+'.join([item.strip() for item in ''.join(name[1] + ' ' + name[0]).split(' ') if item])
		#name_str = name[1].strip() + '+' + name[0].strip()

		ws.write(row_idx, 0, name[1])
		ws.write(row_idx, 1, name[0])

		# Get student email
		s = findHTML(name_str)[0].contents
		if '@' not in s[0]:
			print ''
			print 'Error: ' + name[1] + ' ' + name[0]
			print ''
		else:
			# add to excel
			# print s[0]
			ws.write(row_idx, 2, s)
new_wb.save('testingggg.xls')
