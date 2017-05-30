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

Email Configure
	http://www.anujgakhar.com/2011/12/09/using-macosx-lion-command-line-mail-with-gmail-as-smtp/
	vi /etc/postfix/main.cf
	relayhost = [mail.sssup.it]:465
	smtp_sasl_auth_enable = yes
	smtp_sasl_password_maps = hash:/etc/postfix/sasl_passwd
	smtp_sasl_security_options = noanonymous
	smtp_use_tls = yes

	/etc/postfix/sasl_passwd
	[mail.sssup.it]:465 e.ruffaldi@sssup.it:2MinuRomeo

	sudo chmod 600 /etc/postfix/sasl_passwd
	sudo postmap /etc/postfix/sasl_passwd
	sudo launchctl stop org.postfix.master
	sudo launchctl start org.postfix.master


"""
from termcolor import colored
import urllib2
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

def splitter(data, pred):
    yes, no = [], []
    for d in data:
        (yes if pred(d) else no).append(d)
    return (yes, no)

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
			if z["Status"] is None:
				z["Status"] = ""
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
				
def coloredhtml(t,c):
	return "<span style='color:%s'>%s</span>" % (c,t)
def generate(args,tenowait,tewait,mode):
	if mode == "colored":
		xcolored = colored
		tableformat = ""
	elif mode == "html":
		xcolored = coloredhtml
		tableformat = "html"
	else:
		xcolored = lambda x,y: x
		tableformat = "plain"

	xinf = float("inf")

	for x in tenowait:
		if x["Days Left"] <= 0:
			x["Days Left"] = xcolored(str(x["Days Left"]),"red")
		del x["Days Since"]
	for x in tewait:
		if x["Days Left"] == xinf:
			x["Days Left"] = "(%d)" % x["Days Since"]
		del x["Days Since"]

	body = tabulate.tabulate(tenowait,headers="keys",tablefmt=tableformat)
	if args.web != "":
		body = args.web + "\n\n" + body
	if len(tewait) > 0:
		body += "\n" + "\nWaiting\n" + xcolored(tabulate.tabulate(tewait,headers="keys",tablefmt=tableformat),"cyan")

	return body

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
	parser.add_argument('-w',type=bool,default=True,help="keep wait separated")
	
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
	if args.w:
		wait,nowait = splitter(t,lambda x: str(x["Status"]).lower().find("wait") >= 0)
		t = [x for x in t if x["Days Left"] != -xinf and x["Days Left"] != xinf]
	else:
		wait = []
		nowait = t

	if not args.a:
		t = [x for x in nowait if x["Days Left"] != -xinf and x["Days Left"] != xinf]
	#print tabulate.tabulate(t,headers="keys")

	fields = ["Group","What","Days Left","Days Since"]
	if args.f:
		# first the ones above, then all the others sorted
		fields = fields + sorted(list(cf-set(fields)))

	tewait = [OrderedDict([(f,x[f]) for f in fields]) for x in wait]
	tenowait = [OrderedDict([(f,x[f]) for f in fields]) for x in nowait]

	body = generate(args,tenowait,tewait,"colored")
	print body
	if args.m:
		print "Sending mail to:",args.to,"subject:",args.s
		msg = MIMEText(generate(args,tewait,tenowait,"text"))
		msg['Subject'] = args.s
		if args.to is None:
			raise Exception("missing args.to")
		else:
			msg['From'] = args.to
			msg['To'] = args.to

			if True:
				p = Popen(['mail', '-s',msg["Subject"],msg["To"]], stdout=PIPE, stdin=PIPE, stderr=STDOUT)    
				mail_stdout = p.communicate(input=msg.as_string())[0]
				print mail_stdout
			else:
				s = smtplib.SMTP('localhost')
				s.sendmail(args.to, [args.to], msg.as_string())
				s.quit()
	if args.o:
		webbrowser.open(args.web)
	

if __name__ == '__main__':
	main()