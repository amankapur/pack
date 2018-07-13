import math

from pack.utils import group_by, sort_by


class Paper():

	# store is a pack.store instance
	# which is distributed centerwise
	# subjectwise.
	def __init__(self, store):
		self.store = store


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
	def create_packing(self, packing_levels):
		for p_level in packing_levels:
			name = p_level['name']
			numbering = p_level['numbering']
			if 'sort_order' in numbering:
				self.store.sort_by(numbering['sort_order'])

			from_number = 1
			if 'start_value' in numbering:
				form_number = p_level['numbering']['start_value']

			reset_val = None

			for k in self.store.data:
				d = {}
				if 'reset' in numbering and reset_val != numbering['reset']:
					v = numbering['reset_value']
					from_number = v(k) if callable(v) else v
					reset_val = numbering['reset']

				d['%s_from'%name] = from_number

				r = p_level['ratio']
				ratio = r(k) if callable(r) else r

				q = int(math.ceil(float(k['quantity'])/ratio))

				to_number = from_number + q - 1

				d['%s_count'%name] = q
				d['%s_to'%name] = to_number

				d['quantity'] = k['quantity']
				d['center_number'] = k['center_number']
				from_number = to_number + 1

				if 'packaging' not in k:
					k['packaging'] = {}

				for ke, va in d.iteritems():
					k['packaging'][ke] = va
