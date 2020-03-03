import cPickle as pickle
import math
from pack.utils import group_by, sort_by
from pprint import pprint
from copy import copy

class Store(object):

	def __init__(self, file_path, validator=None, formatter=None, exploder=None, condenser_id=None, condenser=None, is_data=False):
		self.data = []

		self._sort_by = None

		all_data = []

		if is_data:
			all_data = file_path
		else:
			if type(file_path) in [str, unicode]:
				f = open(file_path, "r")
				all_data = pickle.load(f)
				f.close()
			elif isinstance(file_path, list):
				for f_path in file_path:
					f = open(f_path, "r")
					all_data += pickle.load(f)
					f.close()
			else:
				raise "File path must be string or list"


		self.validator = validator
		self.formatter = formatter
		self.exploder = exploder

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

		if condenser_id and condenser and callable(condenser):
			condensed_data = []
			for k, v in group_by(self.data, condenser_id).iteritems():
				condensed_data += condenser(v)
			self.data = condensed_data


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

	def where(self, dikt):
		for d in self.data:
			all_match = True
			for k,v in dikt.iteritems():
				if d[k] != v:
					all_match = False

			if all_match:
				return d
		return None

	def remove(self, arr_dict):
		new_data = []
		removed = []
		for d in self.data:
			arr_values = []
			for dictt in arr_dict:
				values = []
				for k,v in dictt.iteritems():
					values.append(d[k] == v)

				arr_values.append(all(values))

			if any(arr_values):
				removed.append(d)
			else:
				new_data.append(d)

		self.data = new_data
		return removed


	"""
	packing_level is array of :
		{
			'level': 3,
			'name': 'trunk',
			'ratio': lambda d: 300 if d['subject_code']=='E' else 600,
			'numbering': {
				'sort_order': ['set_number', 'subject_code', 'batch_number', 'center_number'],
				'reset': 'set_number',
				'reset_value': 1 or ((x) => Math.ceil(x, 100) +1),
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

				# if 'reset' in numbering:
				# 	if type(numbering['reset'] == list):
				# 		for k in reversed(numbering['reset']):


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

				plural = name + 's'
				if name == 'box':
					plural = 'boxes'
				d['packaging'][plural] = arr


	@staticmethod
	def get_external_packaging(data, max_size, size_key, names=("container", "inner_container"), start_number=1, jump_back=True, group_key=None):

		def _get_size(d):
			if type(size_key) in [str, unicode] and size_key in d:
				return d[size_key]
			elif callable(size_key):
				return size_key(d)

			pprint(d)
			print size_key, type(size_key)
			return None

		number = start_number
		containers = []
		container_size = 0
		current_container = []

		outer_name, inner_name = names
		size_name = outer_name+"_size"
		number_name = outer_name+'_number'

		def base_packing(base_data, number):
			containers = []
			container_size = 0
			current_container = []

			for d in base_data:
				size = _get_size(d)

				# if current container has space add it
				if size + container_size <= max_size:
					current_container.append(d)
					container_size += size
				else:
					# try to fit in some other container for same center
					added = False

					if jump_back:
						for c in containers:
							if size + c[size_name] <= max_size:
								c[inner_name].append(d)
								c[size_name] += size
								added = True
								break
					# alas, create new container
					if not added:
						containers.append({
							number_name: number,
							size_name: container_size,
							inner_name: current_container
						})
						number += 1
						container_size = _get_size(d)
						current_container = [d]

			return containers, current_container, container_size, number

		if not group_key:
			containers, current_container, container_size, number = base_packing(data, number)


		else:
			for g_key, g_data in sorted(group_by(data, group_key).iteritems()):
				g_size = sum([_get_size(d) for d in g_data])

				if g_size > max_size:
					containers, current_container, container_size, number = base_packing(g_data, number)
					# continue

				if g_size + container_size <= max_size:
					current_container += g_data
					container_size += g_size
				else:
					containers.append({
						number_name: number,
						size_name: container_size,
						inner_name: current_container
					})
					number += 1
					container_size = g_size
					current_container = copy(g_data)


		if len(current_container) > 0:
			containers.append({
				number_name: number,
				size_name: container_size,
				inner_name: current_container
			})
			number += 1
			container_size = 0
			current_container = []

		return containers
