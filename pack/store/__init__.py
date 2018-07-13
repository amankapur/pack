import pickle
import math

from pack.utils import group_by, sort_by

class Store(object):

	def __init__(self, file_path, validator=None, formatter=None, exploder=None):
		self.data = []

		self._sort_by = None

		with open(file_path, 'r') as f:

			self.validator = validator
			self.formatter = formatter
			self.exploder = exploder

			all_data = pickle.load(f)

			if exploder and callable(exploder):
				exploded_data = []
				for d in all_data:
					exploded_data += exploder(d)
				all_data = exploded_data

			if validator and callable(validator):
				for d in all_data:
					if validator(d):
						self.data.append(d)
			else:
				self.data = all_data

			if formatter:
				self.data = [formatter(d) for d in self.data]


	def get_total(self, k):
		total = 0
		for d in self.data:
			total += d['k']
		return total

	def sort_by(self, col_names):
		if not self._sort_by == col_names:
			self.data = sort_by(self.data, col_names)
			self._sort_by = col_names
		return self.data
		

	def pluck(self, col_name):
		vals = []
		for d in self.data:
			vals.append(d[col_name])
		return vals

	def get(self, k, v, msg=None):
		for d in self.data:
			if d[k] == v:
				return d

		if msg:
			print msg
		return None

	def update(self, props, v):
		data = []
		for d in self.data:
			arr = []
			for key, val in props.iteritems():
				arr.append(d[key] == val)

			if all(arr):
				new_d = copy(d)
				for key, val in v.iteritems():
					new_d[key] = val
				data.append(new_d)
			else:
				data.append(d)

		self.data = data

	def group_by(self, keys, expect_single=False):
		return group_by(self.data, keys, expect_single)

	"""
	packing_level is array of :
		{
			'level': 3,
			'name': 'trunk',
			'ratio': lambda d: 300 if d['subject_code']=='E' else 600,
			'numbering': {
				'sort_order': ['set_number', 'subject_code', 'batch_number', 'center_number']
			}
		}
	"""
	def create_packaging(self, packing_levels):

		def _get_ratio_val(p_level):
			r = p_level['ratio']
			return r(d) if callable(r) else r

		def _get_splits(p_level, next_name, d):
			splits = None
			if 'custom_splitting' in p_level:
				get_split = p_level['custom_splitting']
				splits = get_split(d['packaging']['%s_count'%next_name], _get_ratio_val(p_level), d)
			return splits

		for p_idx, p_level in enumerate(packing_levels):
			name = p_level['name']
			numbering = {}
			if 'numbering' in p_level:
				numbering = p_level['numbering']
			if 'sort_order' in numbering:
				self.sort_by(numbering['sort_order'])

			from_number = 1
			if 'start_value' in numbering:
				from_number = p_level['numbering']['start_value']

			reset_val = None

			prev_name = None
			if p_idx != 0:
				prev_level = packing_levels[p_idx-1]
				prev_name = prev_level['name']

			for d in self.data:
				if 'packaging' not in d:
					d['packaging'] = {}

				if 'reset' in numbering and reset_val != d[numbering['reset']]:
					v = numbering['reset_value']
					if from_number not in [1, p_level['numbering']['start_value'] if 'start_value' in p_level['numbering'] else 0]:
						from_number = v(from_number, d) if callable(v) else v
					reset_val = d[numbering['reset']]

				d['packaging']['%s_from'%name] = from_number

				ratio = _get_ratio_val(p_level)

				splits = None
				if prev_name:
					splits = _get_splits(p_level, prev_name, d)

				q = int(math.ceil(float(d['quantity'])/ratio))
				if splits and len(splits)>0:
					q = len(splits)

				to_number = from_number + q - 1

				d['packaging']['%s_count'%name] = q
				d['packaging']['%s_to'%name] = to_number

				from_number = to_number + 1

		levels = sort_by(packing_levels, ['-level'])

		for d in self.data:
			p = d['packaging']
			for i, p_level in enumerate(levels):

				if p_level['level'] == 1:
					continue

				name = p_level['name']

				next_level = levels[i+1]
				next_name = next_level['name']
				n_p_from = p['%s_from'%next_name]

				splits = _get_splits(p_level, next_name, d)

				arr = []

				for idx, i in enumerate(range(p['%s_from'%name], p['%s_to'%name]+1)):
					n_p_to = min(n_p_from + _get_ratio_val(p_level)/_get_ratio_val(next_level) - 1, p['%s_to'%next_name])

					count = n_p_to - n_p_from + 1
					if splits:
						count = splits[idx]
						n_p_to = n_p_from + count - 1

					arr.append({
						'%s_number'%name: i,
						'%s_from'%next_name: n_p_from,
						'%s_to'%next_name: n_p_to,
						'%s_count'%next_name: n_p_to - n_p_from + 1
					})
					n_p_from = n_p_to + 1

				d['packaging'][name+'s'] = arr
