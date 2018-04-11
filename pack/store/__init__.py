import pickle

class Store(object):

	def __init__(self, file_path, validator=None, formatter=None, exploder=None):
		self.data = []
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

		self.data = multikeysort(self.data, col_names)
		return self.data

	def pluck(self, col_name):
		vals = []
		for d in self.data:
			vals.append(d[col_name])
		return vals

	def get(self, k, v):
		for d in self.data:
			if d[k] == v:
				return d
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
