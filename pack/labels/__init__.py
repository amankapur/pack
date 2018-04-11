import pdfkit
import os
import string

from pack.utils import replace_all

my_path = os.path.abspath(os.path.dirname(__file__))
base_html = open(my_path + "/view/base.html", 'r').read()
page_html = open(my_path + "/view/page.html", 'r').read()


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
		options[k] = v
	pdfkit.from_string(html, fname, options=options)

def make_labels(data, label_template):
	pages = ""
	for d in data:
		page_data = {
			"#label": ""
		}
		r_data = {}
		for k,v in d.iteritems():
			r_data['#'+str(k)] = str(v)
		page_data["#label"] = replace_all(label_template, r_data)

		pages += replace_all(page_html, page_data)

	render_html = base_html.replace("#pages", pages)
	print "Starting pdf creation : %s " % out_file
	g_pdf(out_file, render_html)
