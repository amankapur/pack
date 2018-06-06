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

		if not self._sort_by == col_names:
			self.data = multikeysort(self.data, col_names)
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
			return _group(self.data, keys)

		data = _group(self.data, keys[0])
		for k, datum in data.iteritems():
			data[k] = _group(datum, keys[1])

		return data

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
		for p_level in packing_levels:
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

			for d in self.data:
				if 'packaging' not in d:
					d['packaging'] = {}

				if 'reset' in numbering and reset_val != numbering['reset']:
					v = numbering['reset_value']
					from_number = v(d) if callable(v) else v
					reset_val = numbering['reset']

				d['packaging']['%s_from'%name] = from_number

				r = p_level['ratio']
				ratio = r(d) if callable(r) else r

				q = int(math.ceil(float(d['quantity'])/ratio))

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

				splits = None
				if 'custom_splitting' in p_level:
					get_split = p_level['custom_splitting']
					splits = get_split(p['%s_count'%next_name])

					if len(range(p['%s_from'%name], p['%s_to'%name]+1)) != len(splits):
						print 'Splits length is not the same for %s : %s' % (name, str(p['%s_count'%next_name]))

				arr = []

				for idx, i in enumerate(range(p['%s_from'%name], p['%s_to'%name]+1)):
					n_p_to = min(n_p_from + p_level['ratio']/next_level['ratio'] - 1, p['%s_to'%next_name])

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
