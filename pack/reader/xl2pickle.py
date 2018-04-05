import os
import pickle
import logging

from openpyxl import load_workbook


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class xl2pickle(object):

	def __init__(self, in_path, out_path, columns, use_all_sheets=False, header_rows=1):

		if not os.path.isdir(in_path):
			raise ValueError('Input path is not a valid directory')

		if not os.path.isdir(out_path):
			raise ValueError('Output path is not a valid directory')

		self.in_path = in_path
		self.out_path = out_path
		self.columns = columns
		self.use_all_sheets = use_all_sheets
		self.header_rows = header_rows

	def _get_formatter(self, t):
		if t == int:
			def f(value):
				return int(float(str(value)))
			return f
		if t == float:
			def f(value):
				return float(str(value))
			return f

	def read_xlsx(self, save=True):
		for fn in os.listdir(self.in_path):
			f = self.in_path+fn
			if os.path.isfile(f) and f.endswith('.xlsx') and not fn.startswith('~'):
				logger.info('Reading from %s : %s' % (self.in_path, fn))
				data = []
				wb = load_workbook(f, data_only=True)
				for sh_name in wb.sheetnames:
					ws = wb[sh_name]
					for r_i in range(self.header_rows+1, ws.max_row+1):
						row_data = {}
						for col in self.columns:
							cell = ws.cell(row=r_i, column=col['index'])
							val = cell.value
							if cell.value != None:
								if 'formatter' in col.keys():
									val = col['formatter'](cell.value)
								elif 'type' in col.keys():
									val = self._get_formatter(col['type'])(cell.value)
								elif type(cell.value) in [str, unicode]:
									val = cell.value.strip()

							row_data[col['name']] = val
						data.append(row_data)
					if not self.use_all_sheets:
						break
				if save:
					self._write_data(fn, data)


	def _write_data(self, fname, data):
		p = self.out_path+'%s.p' % fname.split('.')[0]
		logger.info('Writing to : ' + p)
		with open(p, 'wb') as f:
			pickle.dump(data, f)
