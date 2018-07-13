import datetime
from operator import itemgetter
from pprint import pprint
from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, NamedStyle
from openpyxl.worksheet.pagebreak import Break

from pack.utils import get_str


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


class Sheet():

	def __init__(self, out_file, options={}):
		self.data = []
		self.headers = []
		self.sort_order = []
		self.out_file = out_file

		self.options = {
			'xls': False,
			'title': None,
			'autosize': True,
			'freeze_headers': True
		}

		self._wb = Workbook()
		self._ws = self._wb.active
		if self.options['title']:
			self._ws.title = self.options['title']

		for k,v in options.iteritems():
			self.options[k] = v

		self._max_width = {}
		self._sub_total_idx = None
		self._sub_total_funcs = {}
		self._sub_total_data = {}

		self._grand_total_props = None
		self._group_cols = {}

	def _set_cell(self, value, row_idx, col_idx, style="TEXT"):
		cell = self._ws.cell(row=row_idx, column=col_idx+1, value=value or "")

		if isinstance(value, datetime.datetime) or isinstance(value, datetime.date):
			cell.style = CELL_STYLES['DATE_STYLE']

		styles = CELL_STYLES[style]

		if "alignment" in styles.keys():
			cell.alignment = styles["alignment"]
		if "font" in styles.keys():
			cell.font = styles["font"]
		if 'border' in styles.keys():
			cell.border = styles['border']

		width = len(str(value))
		if col_idx not in self._max_width.keys():
			self._max_width[col_idx] = width
		else:
			self._max_width[col_idx] = max(self._max_width[col_idx], width)

		return cell

	def _insert_row(self, row_data, row_idx, style):
		for col_idx, cell_val in enumerate(row_data):
			self._set_cell(cell_val, row_idx, col_idx, style)

	def _update_sub_total_data(self, row_data):
		if self._sub_total_idx != None:
			for sub_idx, sub_func in self._sub_total_funcs.iteritems():
				if sub_idx not in self._sub_total_data.keys():
					self._sub_total_data[sub_idx] = 0
					if sub_func == 'count_uniq':
						self._sub_total_data[sub_idx] = []

				val = row_data[sub_idx]
				if sub_func == 'count':
					self._sub_total_data[sub_idx] += 1
				elif sub_func == 'sum':
					self._sub_total_data[sub_idx] += val
				elif sub_func == 'min':
					self._sub_total_data[sub_idx] = min(self._sub_total_data[sub_idx], val)
				elif sub_func == 'max':
					self._sub_total_data[sub_idx] = max(self._sub_total_data[sub_idx], val)
				elif sub_func == 'count_uniq':
					self._sub_total_data[sub_idx] = list(set(self._sub_total_data[sub_idx]+[val]))

				self._sub_total_data[self._sub_total_idx] = row_data[self._sub_total_idx]

	def _insert_subtotal_row(self, row_idx, page_break=True):
		r = []
		for i in range(len(self.headers[-1])):
			r.append("")

		for idx, val in self._sub_total_data.iteritems():
			if idx == self._sub_total_idx:
				r[idx] = get_str(val) + ' Total'
			else:
				if type(val) == list:
					r[idx] = len(val)
				else:
					r[idx] = val

		self._insert_row(r, row_idx, 'HEADER')
		if page_break:
			self._ws.page_breaks.append(Break(id=row_idx))
		self._sub_total_data = {}

	def _insert_grandtotal_row(self, row_idx):
		r = []
		for i in self.headers[-1]:
			r.append("")

		r[1] = 'Grand Total'

		sorted_data = self._get_sorted_data()
		for col_idx, grand_total_func in self._grand_total_props.iteritems():
			val_array = [d[col_idx] for d in sorted_data]
			if grand_total_func == 'sum':
				val = sum(val_array)
			elif grand_total_func == 'min':
				val = min(val_array)
			elif grand_total_func == 'max':
				val = max(val_array)
			elif grand_total_func == 'count':
				val = len(val_array)

			r[col_idx] = val

		self._insert_row(r, row_idx, 'TITLE')


	def _get_sorted_data(self):
		sorted_data = self.data

		if len(self.sort_order) > 0:
			for i in reversed(self.sort_order):
				v = i
				reverse = False
				if i < 0:
					v = -1 * i
					reverse = True
				sorted_data.sort(key=itemgetter(v), reverse=reverse)

		return sorted_data

	def _insert_headers(self):
		r_idx = 1
		for h in self.headers:
			self._insert_row(h, r_idx, 'HEADER')
			r_idx += 1
		return r_idx

	def _insert_data(self, r_idx):
		sub_total_val = None
		sorted_data = self._get_sorted_data()
		has_subtotal = self._sub_total_idx != None
		for i, row_data in enumerate(sorted_data):
			if has_subtotal:
				if sub_total_val == None:
					sub_total_val = row_data[self._sub_total_idx]
				if sub_total_val != row_data[self._sub_total_idx]:
					self._insert_subtotal_row(r_idx)
					r_idx += 1
					sub_total_val = row_data[self._sub_total_idx]

			if i != 0:
				prev_row_data = sorted_data[i-1]
				for group_col_idx, group_col_val in self._group_cols.iteritems():
					if row_data[group_col_idx] == group_col_val:
						row_data[group_col_idx] = None
					else:
						self._group_cols[group_col_idx] = row_data[group_col_idx]
			else:
				for group_col_idx, group_col_val in self._group_cols.iteritems():
					self._group_cols[group_col_idx] = row_data[group_col_idx]

			self._insert_row(row_data, r_idx, 'TEXT')

			if has_subtotal:
				self._update_sub_total_data(row_data)
				if i == len(sorted_data)-1:
					r_idx += 1
					self._insert_subtotal_row(r_idx, False)
					r_idx += 1

			r_idx += 1

		if self._grand_total_props != None:
			self._insert_grandtotal_row(r_idx)
			r_idx += 1
		return r_idx

	def _set_print_headers(self):
		if len(self.headers) > 0:
			self._ws.print_title_rows = '1:%s' % str(len(self.headers))

	def _auto_size_columns(self):
		if self.options['autosize']:
			for col_idx, max_w in self._max_width.iteritems():
				self._ws.column_dimensions[chr(col_idx+97).upper()].width = max_w

	def _freeze_headers(self):
		if self.options['freeze_headers']:
			self._ws.freeze_panes = "A%s" % str(len(self.headers)+1)


	def _create(self):
		print 'Creating sheet  %s ' % self.out_file
		self._wb.save(self.out_file)

	# needs ssconvert installed
	def _xls(self):
		if self.options['xls']:
			p = self.out_file
			os.system("ssconvert %s %s" % (p, p[:-1]))


	def add_headers(self, headers):
		self.headers = headers[0:-1]

		last_header = headers[-1]
		last_header_text = []

		s_hsh = {}

		for col_idx,h in enumerate(last_header):
			if type(h) == tuple:
				text, d = h
				last_header_text.append(text)

				sort_idx = d
				if type(d) != int:
					if 'sort' in d:
						sort_idx = d['sort']
					else:
						sort_idx=None

				if sort_idx:
					s_hsh[sort_idx] = col_idx

				if type(d) == dict:
					if 'sub_total' in d:
						self._sub_total_idx = col_idx

					if 'sub_total_func' in d:
						self._sub_total_funcs[col_idx] = d['sub_total_func']

					if 'grand_total_func' in d:
						if self._grand_total_props == None:
							self._grand_total_props = {}
						self._grand_total_props[col_idx] = d['grand_total_func']

					if 'group_col' in d:
						self._group_cols[col_idx] = None

			else:
				last_header_text.append(h)

		for i in range(len(last_header)+1):
			if i in s_hsh.keys():
				self.sort_order.append(s_hsh[i])
			elif -1*i in s_hsh.keys():
				self.sort_order.append(-1*s_hsh[-1*i])

		self.headers.append(last_header_text)


	def add_rows(self, rows):
		self.data = rows

	def save(self):
		r_idx = self._insert_headers()
		self._insert_data(r_idx)

		self._set_print_headers()
		self._auto_size_columns()
		self._freeze_headers()
		self._create()
		self._xls()
