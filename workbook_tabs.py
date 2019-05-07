import sys, os, time as t
import random
start_clock = t.time()
start_date = t.strftime('%Y-%m-%d %H:%M:%S')

rec = ['file', 'server', 'user', 'start_clock', 'stop_clock', 'elapsed_seconds', 'stop_date']
colors = ['red','green','blue','orange','yellow']

file_name = 'Dynamic Workbook.xlsx'

import xlsxwriter

# Dynamic Construction
workbook = xlsxwriter.Workbook(file_name)

w = 'worksheet'
n = 0

for x in rec[1:]:
	n+=1
	s = w+str(n)
	workbook.add_worksheet().set_tab_color(random.choice(colors))

workbook.close()

# Direct Construction
file_name = 'Direct Workbook.xlsx'
workbook = xlsxwriter.Workbook(file_name)

sheet1 = workbook.add_worksheet().set_tab_color('red')
sheet2 = workbook.add_worksheet().set_tab_color('white')
sheet3 = workbook.add_worksheet().set_tab_color('blue')

workbook.close()