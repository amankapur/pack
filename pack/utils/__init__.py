import math
import datetime

from mailmerge import MailMerge
from copy import copy

from openpyxl.styles import Alignment, Font, Border, Side, NamedStyle

from pprint import pprint

import os


# needs ssconvert installed
def convert_to_xls(in_path):
	os.system("ssconvert %s %s" % (in_path, in_path[:-1]))

def get_min(arr_of_dicts, key):
	min_val = float("inf")

	for d in arr_of_dicts:
		min_val = min(min_val, d[key])
	return min_val

def get_max(arr_of_dicts, key):
	max_val = float("-inf")

	for d in arr_of_dicts:
		max_val = max(max_val, d[key])
	return max_val


#keymap is a has str->str
# it maps keys of the in_doc merge fields to data keys/funcs
# data is [{..},{...}] format
def my_merge(in_doc_path, out_doc_path, data, key_map=None):
	mm_data = []
	if key_map:
		for d in data:
			mm_d = {}

			for mm_key, d_key in key_map.iteritems():
				v = ''
				if callable(d_key):
					v = d_key(d)
				else:
					v = d[d_key]

				mm_d[mm_key] = v

			mm_data.append(mm_d)
	else:
		mm_data = data

	formatted_d = [{},{}]
	for d in mm_data:
		for k,v in d.iteritems():
			d[k] = str(v)
		formatted_d.append(d)

	MAX_SIZE_PER_FILE = 5000
	with MailMerge(in_doc_path) as document:
		document.merge_pages(formatted_d[:MAX_SIZE_PER_FILE])
		document.write(out_doc_path)


CELL_STYLES = {
	'HEADER':{
		'alignment': Alignment(horizontal='center', vertical='center', wrap_text="true"),
		'font': Font(bold=True),
		'border':  Border(left=Side(style='thin'),
											 right=Side(style='thin'),
											 top=Side(style='thin'),
											 bottom=Side(style='thin'))
	},
	'TEXT':{
		'alignment': Alignment(horizontal='center', vertical='center'),
		'border':  Border(left=Side(style='thin'),
											 right=Side(style='thin'),
											 top=Side(style='thin'),
											 bottom=Side(style='thin'))
	},
	'TITLE':{
		'alignment': Alignment(horizontal='center', vertical='center'),
		'font': Font(bold=True, size=15)
	},
	'DATE_STYLE': NamedStyle(name='datetime', number_format='DD/MM/YYYY')
}

def insert_row(sheet, data, row_num, styles={}):
	col_idx = 1
	for d in data:
		cell = sheet.cell(row=row_num, column=col_idx, value=d or "")

		if isinstance(d, datetime.datetime) or isinstance(d, datetime.date):
			cell.style = CELL_STYLES['DATE_STYLE']

		if "alignment" in styles.keys():
			cell.alignment = styles["alignment"]
		if "font" in styles.keys():
			cell.font = styles["font"]
		if 'border' in styles.keys():
			cell.border = styles['border']



		col_idx +=1


def copy_style(new_cell, cell):
	new_cell.font = copy(cell.font)
	new_cell.border = copy(cell.border)
	new_cell.fill = copy(cell.fill)
	new_cell.number_format = copy(cell.number_format)
	new_cell.protection = copy(cell.protection)
	new_cell.alignment = copy(cell.alignment)

def get_str(k):
	if type(k) == datetime.datetime:
		d = str(k.day)
		if len(d) < 2:
			d = '0' + d

		m = str(k.month)
		if len(m) < 2:
			m = '0' + m

		y = str(k.year)
		if len(y) > 2:
			y = y[2:]

		return '%s/%s/%s'%(d,m,y)
	else:
		return str(k)

def display_from_to(from_val, to_val):
	if from_val is None or to_val is None:
		return None
	if from_val == to_val:
		return str(from_val)
	else:
		return str(from_val) + ' - ' + str(to_val)

def display_time(t):
	a = ' AM'
	if t.hour > 12:
		a = ' PM'
	return str(t.hour%12) + ':' + str(t.minute) + a

def display_month(m):
	if 'MAR' in m :
		return 'MARCH'
	return m

def round_up(x, n):
	n = float(n)
	return int(math.ceil(x / n) * n)

def get_string_sub_code(code):
	c = str(code)
	if len(c) != 2:
		c = '0'+ c
	return c

def zerofy(n):
	if n < 10:
		return '0' + str(n)
	else:
		return str(n)

def time_plus(time, timedelta):
	start = datetime.datetime(
			2000, 1, 1,
			hour=time.hour, minute=time.minute, second=time.second)
	end = start + timedelta
	return end.time()



def split_arr_n(arr, x):
	n = int(math.ceil(len(arr)/float(x)))

	new_arr = []
	for i in range(0, n):
		for j in range(0, x):
			new_index = i+j*n
			if new_index < len(arr):
				new_arr.append(arr[new_index])
			# else:
			#   new_arr.append(None)

	return new_arr

def sort_by(items, columns):
	return multikeysort(items, columns)

def multikeysort(items, columns):
	from operator import itemgetter
	comparers = [ ((itemgetter(col[1:].strip()), -1) if col.startswith('-') else (itemgetter(col.strip()), 1)) for col in columns]
	def comparer(left, right):
			for fn, mult in comparers:
					result = cmp(fn(left), fn(right))
					if result:
							return mult * result
			else:
					return 0
	return sorted(items, cmp=comparer)

def group_by(input_data, keys, expect_single=False):
	def _group(datum, k):
		data = {}
		for d in datum:
			if d[k]:
				if d[k] not in data.keys():
					if expect_single:
						data[d[k]] = d
					else:
						data[d[k]] = [d]
				elif expect_single:
					raise ValueError('Expected single item in group_by')
				else:
					data[d[k]].append(d)
		return data

	if type(keys) in [str, unicode]:
		return _group(input_data, keys)

	data = _group(input_data, keys[0])
	for k, datum in data.iteritems():
		data[k] = _group(datum, keys[1])

	return data


def replace_all(string, d):
	s = string
	for k,v in d.items():
		s = s.replace(k, str(v))
	return s
