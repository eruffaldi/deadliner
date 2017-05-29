"""Deadliner with Web storage, Emanuele Ruffaldi 2017

Input: Excel XLSX or Web link

Output: 
	open web page (-o)
	tabulate (none)
	send email (-m)

Model:
	tasks divided in groups
	each task has
		Start Date
		Deadline Date
	optional:
		Effective Date: when completed reports the effective due date
	computed in Excel:
		Days Left: can be empty for open ended tasks or unknown deadline (+inf)
		Days Since: tracks amount time since start (inf)

Entities are ordered by (-DaysLeft +DaysSince), we could add group priority

"""
import urllib2
import argparse
from openpyxl import *
import os
import tabulate
import pprint
from collections import OrderedDict
import smtplib
from email.mime.text import MIMEText
from subprocess import Popen, PIPE, STDOUT
import argparse
import webbrowser
import cStringIO


def findregions(ws):
	#Type	Status	Started	Due	Days Left	Days Since
	lasthead = ""
	inpart = False
	rows = 0
	regions = dict()
	for i,r in enumerate(ws.values):
		if not inpart:
			if r[1] == "Type" and r[2] == "Status":
				fields = [x for x in r if x != "" and x is not None]
				regionname = lasthead
				#print "found",lasthead,fields
				inpart = True
				rows = 0
				relrow = i+1
			else:
				lasthead = r[0]
		else:
			if r[0] is None:
				#print "ended with",rows
				inpart = False
				lasthead = ""
				regions[regionname] = dict(rows=rows,endrow=i,fields=fields,startrow=relrow)
			else:
				rows += 1
	return regions

def findtasks(ws,q):
	allfields = reduce(lambda x,y: x|y,[set(x["fields"]) for x in q.values()],set())
	allfields.add("Group")
	template = dict([(x,"") for x in allfields])
	removefields = set(["Class","Location","Effective","Head"]) & allfields

	t = []
	# any group
	xinf = float("inf")
	for g,y in q.iteritems():
		ff = y["fields"]
		# any entry by row
		for i in range(y["startrow"],y["endrow"]):
			z = {}
			z.update(template)
			z["Group"] = g
			for j,f in enumerate(ff):
				z[f] = ws.cell(row=i+1,column=j+1).value
			if z["Days Left"] is None or z["Days Left"] == "":
				z["Days Left"] = xinf
			else:
				z["Days Left"] = float(z["Days Left"])
			if z["Days Since"] is None or z["Days Since"] == "":
				z["Days Since"] = -xinf
			else:
				z["Days Since"] = float(z["Days Since"])
			for r in removefields:
				del z[r]
			t.append(z)
	return (t,allfields-removefields)
				
def main():

	parser = argparse.ArgumentParser(description='Deadliner')
	parser.add_argument('-m',action="store_true",help="send email")
	parser.add_argument('-o',action="store_true",help="open web link")
	parser.add_argument('--to',help="send email target")
	parser.add_argument('-i',help="input link or file")
	parser.add_argument('-a',action="store_true",help="all actions, otherwise remove open ended")
	parser.add_argument('--web',help="open link")
	parser.add_argument('-s',help="Subject",default="Task Deadlines")
	parser.add_argument('-f',help="Full Fields",action="store_true")
	
	args = parser.parse_args()

	if args.i.startswith("http"):
		response = urllib2.urlopen(args.i)
		#open(tmp,"wb").write()
		tmp = cStringIO.StringIO(response.read())
	else:
		tmp = args.i

	wb = load_workbook(tmp,data_only=True)
	q = findregions(wb["Tasks"])
	t,cf = findtasks(wb["Tasks"],q)
	t.sort(key=lambda x: (x["Days Left"],-x["Days Since"]))
	xinf = float("inf")
	# remove not actiond
	if args.a:
		t = t
	else:
		t = [x for x in t if x["Days Left"] != -xinf and x["Days Left"] != xinf]
	#print tabulate.tabulate(t,headers="keys")

	fields = ["Group","What","Days Left"]
	if args.f:
		# first the ones above, then all the others sorted
		fields = fields + sorted(list(cf-set(fields)))
	te = [OrderedDict([(f,x[f]) for f in fields]) for x in t] 
	body = tabulate.tabulate(te,headers="keys")
	body = args.web + "\n" + body
	print body
	if args.m:
		print "Sending mail to:",args.to,"subject:",args.s
		msg = MIMEText(body)
		msg['Subject'] = args.s
		if args.to is None:
			raise Exception("missing args.to")
		else:
			msg['From'] = args.to
			msg['To'] = args.to

			p = Popen(['mail', '-s',msg["Subject"],msg["To"]], stdout=PIPE, stdin=PIPE, stderr=STDOUT)    
			mail_stdout = p.communicate(input=body)[0]
			print mail_stdout
			# Send the message via our own SMTP server, but don't include the
			# envelope header.
			#s = smtplib.SMTP('localhost')
			#s.sendmail(me, [you], msg.as_string())
			#s.quit()
	if args.o:
		webbrowser.open(args.web)
	

if __name__ == '__main__':
	main()