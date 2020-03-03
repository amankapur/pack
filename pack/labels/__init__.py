import pdfkit
import os
import string
import math

from pprint import pprint

from pack.utils import replace_all, get_str
from datetime import datetime
from PyPDF2 import PdfFileReader


def g_pdf(fname, html, custom_options={}):

	# MUST ADD NEW PROPS HERE ELSE NOT USED
	options = {
		'page-width': 165.4,
		'page-height': 165.4,
		'disable-smart-shrinking': None,
		'margin-bottom': '0',
		'margin-top': '0',
		'margin-left': '0',
		'margin-right': '0',
		'print-media-type': '',
		'quiet': '',
		'dpi': 300,
		'orientation': 'Portrait'
	}

	for k,v in custom_options.iteritems():
		if k in options:
			options[k] = v

	if 'page-size' in custom_options:
		options.pop('page-height', None)
		options.pop('page-width', None)
		options['page-size'] = custom_options['page-size']

	# pprint(options)
	pdfkit.from_string(html, fname, options=options)

def make_labels(label_data, label_template, output_file, options={}, debug=False):
	defaults = {
		'custom_style': None,
		'max_label_per_file': len(label_data)+1,
		'duplicate': False,
		'page-height': 165.4,
		'page-width': 165.4,
		'check_pages': True
	}

	for k,v in defaults.iteritems():
		if k not in options:
			options[k] = v

	data = None
	if options['duplicate']:
		data = []
		for l_d in label_data:
			for i in range(2):
				data.append(l_d)
	else:
		data = label_data

	MAX_FILE_SIZE = options['max_label_per_file']

	file_count = int(math.ceil(len(data)/float(MAX_FILE_SIZE)))

	f_names = []
	for i in range(file_count):
		start = i*MAX_FILE_SIZE
		end = min((i+1)*MAX_FILE_SIZE, len(data))
		out_file = None
		if file_count > 1:
			out_file = output_file.split('.pdf')[0] + '_PART_%s.pdf' % str(i+1)
		else:
			out_file = output_file
		print 'Creating %s labels : %s ' % (str(end-start), out_file)
		f_names.append(out_file)
		_create_html_labels(data[start:end], label_template, out_file, options, debug)
	return f_names


def _create_html_labels(data, label_template, output_file, options, debug):
	my_path = os.path.abspath(os.path.dirname(__file__))
	base_html = open(my_path + "/view/base.html", 'r').read()
	height = str(options['page-height'])  + "mm";
	if 'page-size' in options:
		height = "100%"
	base_html = base_html.replace("#PAGE_HEIGHT", height)
	page_html = open(my_path + "/view/page.html", 'r').read()

	pages = ""
	for d in data:
		page_data = {
			"#label": ""
		}
		r_data = {}
		for k,v in d.iteritems():
			val = v
			if type(v) == datetime:
				val = get_str(v)
			elif type(v) not in [unicode, str]:
				val = str(v or "")
			r_data['#'+str(k)] = val
		
		page_data["#label"] = replace_all(label_template, r_data)

		pages += replace_all(page_html, page_data)

	render_html = base_html.replace("#pages", pages)

	custom_style = ""
	if 'custom_style' in options and options['custom_style']:
		custom_style = options['custom_style']

	render_html = render_html.replace("#custom_style", custom_style)

	if debug:
		print 'Output debug file created debug_html.html'
		debug = open('debug_html.html', 'w')
		debug.write(render_html.encode('utf8'))
		debug.close()

	d_name = "/".join(output_file.split("/")[:-1])
	if not os.path.isdir(d_name):
		os.makedirs(d_name)

	g_pdf(output_file, render_html, options)

	if options['check_pages']:
		created_pages = PdfFileReader(open(output_file, "rb")).getNumPages()
		if created_pages != len(data):
			print 'PAGE MISMATCH : %s created, %s asked' % (str(created_pages), str(len(data)))

def create_id_label(id_name, out_file, options={}):
	my_path = os.path.abspath(os.path.dirname(__file__))
	base_html = open(my_path + "/view/id.html", 'r').read()
	render_html = base_html.replace('#ID_NAME', id_name)
	g_pdf(out_file, render_html, options)
