from docx import *
from docx.shared import Inches
import pprint
import codecs

l = 0
current = 'p'
students = {}
student = {}

def get_grade(s):
	if s >= 70:
		return "First"
	if s >= 67 and s < 70:
		return "High 2:1"
	if s >= 63 and s < 67:
		return "Mid 2:1"
	if s >= 60 and s < 64:
		return "Low 2:1"
	if s >= 57 and s < 60:
		return "High 2:2"
	if s >= 53 and s < 57:
		return "Mid 2:2"
	if s >= 50 and s < 54:
		return "Low 2:2"
	if s >= 47 and s < 50:
		return "High 3rd"
	if s >= 43 and s < 47:
		return "Mid 3rd"
	if s >= 40 and s < 44:
		return "Low 3rd"
	if s >= 35 and s < 40:
		return "Marginal Fail"
	if s < 35:
		return "Fail"

def output_part(doc, p, part_id):
	#doc.add_heading(part_id, 3)
	doc.add_paragraph('- '+part_id, 'List2')
	for pp in p:
		doc.add_paragraph(pp, style='ListBullet3')

def output(doc, tree, show_score):
	if not tree.has_key('name'):
		return
	doc.add_heading(tree['name'], 1) # name
	total_s = 0
	scores = ''
	for part_id in tree['parts']:
		part_name = ''
		if part_id == 1:
			part_name = 'A'
			doc.add_heading('Part-'+str(part_name), 2) # part a
		if part_id == 2:
			part_name = 'B'
			doc.add_heading('Part-'+str(part_name), 2) # part a
		if part_id == 3:
			part_name = 'C'
			doc.add_heading('Part-'+str(part_name), 2) # part a
		if part_id == 4:
			part_name = 'Overall'
			doc.add_heading(str(part_name), 2) # part a
		part = tree['parts'][part_id]
		p = part['p']
		j = part['j']
		s = part['s']
		s = int(s)
		scores = scores + ', ' + str(s)
		total_s = total_s + s
		output_part(doc, p, 'Pseudo code:')
		output_part(doc, j, 'JS code:')
	if (show_score == True):
		doc.add_heading('Scores = '+scores +' = ' + str(total_s), 3)
	grade = get_grade(total_s)
	doc.add_heading('Final grade = '+grade, 3)
	doc.add_page_break()

def output_txt(doc, tree, show_score):
	if not tree.has_key('name'):
		return
	name_id = tree['name']
	print '====', tree['name'], tree, '===='
	name_id = name_id.split('(')
	name = name_id[0].strip()
	id = name_id[1].replace(')', '').strip()
	total_s = 0
	for part_id in tree['parts']:
		part = tree['parts'][part_id]
		s = part['s']
		s = int(s)
		total_s = total_s + s
	grade = get_grade(total_s)
	doc.write('%s\t%s\t%s\t%s\t\n'%(id, name, grade, total_s))

def output_loop(doc, pas, show_score):
	for name in pas:
		pa = pas[name]
		output(doc, pa, show_score)

def output_excel(doc, tree, show_score):
	name = ''
	id = ''
	grade = ''
	part_a = ''
	part_b = ''
	part_c = ''
	overall = ''
	###
	if not tree.has_key('name'):
		return
	##
	name_id = tree['name']
	name_id = name_id.split('(')
	name = name_id[0].strip()
	id = name_id[1].replace(')', '').strip()
	##
	total_s = 0
	scores = ''
	for part_id in tree['parts']:
		part = tree['parts'][part_id]
		p = part['p']
		j = part['j']
		s = part['s']
		s = int(s)
		scores = scores + ', ' + str(s)
		total_s = total_s + s
		part_name = ''
		if part_id == 1:
			part_name = 'A'
			part_a = part
		if part_id == 2:
			part_name = 'B'
			part_b = part
		if part_id == 3:
			part_name = 'C'
			part_c = part
		if part_id == 4:
			part_name = 'Overall'
			part_g = part
	if (show_score == True):
		doc.add_heading('Scores = '+scores +' = ' + str(total_s), 3)
	grade = get_grade(total_s)
	##
	print name, id, total_s, grade, part_a, part_b, part_c, part_g

def check(t, post, doc, show_score):
	global l, pa, current
	#print '******', t, post
	if t == '' and not post == '':
		l = 0
		pprint.pprint(pa)
		print '\n\n\n'
		print 'name====', post
		output(doc, pa, show_score)
		pa = {'parts':{}, 'name': post}
	if 'part a:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 1
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'part b:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 2
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'part c:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 3
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'overall:' in t.lower():
		ss = t.lower().split(':')
		l = 4
		pa['parts'][l] = {'s':0, 'p':[], 'j':[]}
	elif 'pseudo code:' in t.lower():
		current = 'p'
	elif 'js code:' in t.lower():
		current = 'j'
	elif 'all:' in t.lower() or 'you have therefore' in t.lower():
		return
	else:
		#print t
		if l == 0:
			return
		if t == '':
			return
		pa['parts'][l][current].append(t)

def check_txt(t, post, doc, show_score):
	global l, pa, current
	#print '******', t, post
	if t == '' and not post == '':
		l = 0
		pprint.pprint(pa)
		print '\n\n\n'
		print 'name====', post
		output_txt(doc, pa, show_score)
		pa = {'parts':{}, 'name': post}
	if 'part a:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 1
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'part b:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 2
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'part c:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 3
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'overall:' in t.lower():
		ss = t.lower().split(':')
		l = 4
		pa['parts'][l] = {'s':0, 'p':[], 'j':[]}
	elif 'pseudo code:' in t.lower():
		current = 'p'
	elif 'js code:' in t.lower():
		current = 'j'
	elif 'all:' in t.lower() or 'you have therefore' in t.lower():
		return
	else:
		#print t
		if l == 0:
			return
		if t == '':
			return
		pa['parts'][l][current].append(t)

def check_excel(t, post, doc, show_score):
	global l, pa, current
	#print '******', t, post
	if t == '' and not post == '':
		l = 0
		pprint.pprint(pa)
		print '\n\n\n'
		print 'name====', post
		output_excel(doc, pa, show_score)
		pa = {'parts':{}, 'name': post}
	if 'part a:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 1
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'part b:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 2
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'part c:' in t.lower():
		ss = t.lower().split(':')
		s = ss[1]
		l = 3
		pa['parts'][l] = {'s':s, 'p':[], 'j':[]}
	elif 'overall:' in t.lower():
		ss = t.lower().split(':')
		l = 4
		pa['parts'][l] = {'s':0, 'p':[], 'j':[]}
	elif 'pseudo code:' in t.lower():
		current = 'p'
	elif 'js code:' in t.lower():
		current = 'j'
	elif 'all:' in t.lower() or 'you have therefore' in t.lower():
		return
	else:
		#print t
		if l == 0:
			return
		if t == '':
			return
		pa['parts'][l][current].append(t)

def read_doc(doc_name, save_doc_name, show_score):
	doc = Document()
	pre = ''
	document = Document(doc_name)
	for post in document.paragraphs:
		t = post.text.strip()
		check(pre, t, doc, show_score)
		pre = t
	pprint.pprint(pa)
	output(doc, pa, show_score)
	doc.save(save_doc_name)

def read_doc_txt(doc_name, save_doc_name, show_score):
	doc = codecs.open(save_doc_name, 'w', encoding='utf-8')
	pre = ''
	document = Document(doc_name)
	for post in document.paragraphs:
		t = post.text.strip()
		check_txt(pre, t, doc, show_score)
		pre = t
	pprint.pprint(pa)
	output_txt(doc, pa, show_score)
	doc.close()

def read_doc_excel(doc_name, save_doc_name, show_score):
	doc = codecs.open(save_doc_name, 'w', encoding='utf-8')
	pre = ''
	document = Document(doc_name)
	for post in document.paragraphs:
		t = post.text.strip()
		check_excel(pre, t, doc, show_score)
		pre = t
	pprint.pprint(pa)
	output_excel(doc, pa, show_score)
	doc.close()

if __name__ == "__main__":
	'''
	####
	l = 0
	current = 'p'
	pa = {}
	doc_name = './Marking-Group-H.docx'
	save_doc_name = 'H-With-Score.docx'
	show_score = True
	read_doc(doc_name, save_doc_name, show_score)
	save_doc_name = 'H.docx'
	show_score = False
	read_doc(doc_name, save_doc_name, show_score)
	####
	l = 0
	current = 'p'
	pa = {}
	doc_name = './Marking-Group-F.docx'
	save_doc_name = 'F-With-Score.docx'
	show_score = True
	read_doc(doc_name, save_doc_name, show_score)
	save_doc_name = 'F.docx'
	show_score = False
	read_doc(doc_name, save_doc_name, show_score)
	####
	l = 0
	current = 'p'
	pa = {}
	doc_name = './Marking-Group-H.docx'
	save_doc_name = 'H-With-Score.txt'
	show_score = True
	read_doc_txt(doc_name, save_doc_name, show_score)
	####
	l = 0
	current = 'p'
	pa = {}
	doc_name = './Marking-Group-F.docx'
	save_doc_name = 'F-With-Score.txt'
	show_score = True
	read_doc_txt(doc_name, save_doc_name, show_score)
	####
	'''
	####
	l = 0
	current = 'p'
	pa = {}
	doc_name = './Marking-Group-H.docx'
	save_doc_name = 'H-individual.txt'
	show_score = True
	read_doc_excel(doc_name, save_doc_name, show_score)
	####
	l = 0
	current = 'p'
	pa = {}
	doc_name = './Marking-Group-F.docx'
	save_doc_name = 'F-individual.txt'
	show_score = True
	read_doc_excel(doc_name, save_doc_name, show_score)


