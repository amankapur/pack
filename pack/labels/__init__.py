import pdfkit
import os
import string
import math

from pprint import pprint

from pack.utils import replace_all


def g_pdf(fname, html, custom_options={}):

	options = {
		'page-width': 165.4,
		'page-height': 165.4,
		'disable-smart-shrinking': None,
		'margin-bottom': '0',
		'margin-top': '0',
		'margin-left': '0',
		'margin-right': '0',
		'dpi': 300
	}

	for k,v in custom_options.iteritems():
		if k in options:
			options[k] = v
	pdfkit.from_string(html, fname, options=options)

def make_labels(label_data, label_template, output_file, options={}, debug=False):
	defaults = {
		'custom_style': None,
		'max_label_per_file': len(label_data)+1,
		'duplicate': False,
		'page-height': 165.4,
		'page-width': 165.4
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
	for i in range(file_count):
		start = i*MAX_FILE_SIZE
		end = min((i+1)*MAX_FILE_SIZE, len(data))
		out_file = None
		if file_count > 1:
			out_file = output_file.split('.pdf')[0] + '_PART_%s.pdf' % str(i+1)
		else:
			out_file = output_file
		print 'Creating %s labels : %s ' % (str(end-start), out_file)
		_create_html_labels(data[start:end], label_template, out_file, options, debug)


def _create_html_labels(data, label_template, output_file, options, debug):
	my_path = os.path.abspath(os.path.dirname(__file__))
	base_html = open(my_path + "/view/base.html", 'r').read()
	base_html = base_html.replace("#PAGE_HEIGHT", str(options['page-height']))
	page_html = open(my_path + "/view/page.html", 'r').read()

	pages = ""
	for d in data:
		page_data = {
			"#label": ""
		}
		r_data = {}
		for k,v in d.iteritems():
			r_data['#'+str(k)] = str(v or "")
		page_data["#label"] = replace_all(label_template, r_data)

		pages += replace_all(page_html, page_data)

	render_html = base_html.replace("#pages", pages)

	custom_style = ""
	if 'custom_style' in options.keys():
		custom_style = options['custom_style']

	render_html = render_html.replace("#custom_style", custom_style)

	if debug:
		print 'Output debug file created debug_html.html'
		debug = open('debug_html.html', 'w')
		debug.write(render_html)
		debug.close()

	d_name = "/".join(output_file.split("/")[:-1])
	if not os.path.isdir(d_name):
		os.mkdir(d_name)

	g_pdf(output_file, render_html, options)

def create_id_label(id_name, out_file, options={}):
	my_path = os.path.abspath(os.path.dirname(__file__))
	base_html = open(my_path + "/view/id.html", 'r').read()
	render_html = base_html.replace('#ID_NAME', id_name)
	g_pdf(out_file, render_html, options)
