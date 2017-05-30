import datetime

def mergeranges(l):
	if len(l) < 2:
		return l
	l1 = l[0]
	lo = []
	for i in range(1,len(l)):
		li = l[i]
		if li[0] > l1[1]:
			# hole between
			lo.append(l1)
			l1 = li
		elif li[1] <= l1[1]:
			# contained, skip
			pass 
		else:
			# extend
			l1 = (l1[0],li[1])
	lo.append(l1)
	return lo

class BackwardRunner:
	def __init__(self,ranges):
		self.r = ranges
		self.pos = len(self.r)-1
	def isInside(self,value):
		while True:
			# after => no move
			if  value > self.r[self.pos][1]:
				return False
			# inside => no move
			elif value >= self.r[self.pos][0]:
				return True
			# before => move back and loop
			else:
				if self.pos != 0:
					self.pos -= 1
				else:
					# before first
					return False


def computeworkdays(interval,daysoff,workOnSaturdays=False):
	def validatedate(x):
		if type(x) is str:
			return datetime.datetime.strptime(x,"%Y/%m/%d").date()
		elif isinstance(x,datetime.datetime):
			return x.date()
		else:
			return x
	def validaterange(x):
		if not isinstance(x,tuple):
			x = validatedate()
			return (x,x)
		else:
			x,y = x
			x = validatedate(x)
			y = validatedate(y)
			if y < x:
				return (y,x)
			else:
				return (x,y)
	first,last = validaterange(interval) 

	# validate days off
	daysoff = [validaterange(x) for x in daysoff]
	daysoff.sort(key=lambda x:x[0])
	daysoff = mergeranges(daysoff)

	r = 0
	day = last	
	out = []
	wday = day.weekday()
	dt = datetime.timedelta(days=-1)
	br = BackwardRunner(daysoff)
	wdaySaturday = 5 if not workOnSaturdays else 10
	print "start from",day,wday
	print "check saturdays as ",wdaySaturday
	while day > first:
		isworkday = True
		if wday == 6 or wday == wdaySaturday:
			#pass
			isworkday = False
			print day,wday,"sun/sat"
		elif br.isInside(day):
			isworkday = False
			print day,wday,"holiday"
		if isworkday:
			r += 1
		out.append((day,r))

		# back working day and current day
		if wday == 0:
			wday = 6
		else:
			wday -= 1
		day += dt

	print "out"

	return reversed(out)
def main():
	holidays = [
		("2017/06/04","2017/06/07"),
		("2017/06/01","2017/06/02"),
		("2017/07/01","2017/07/10"),
		("2017/12/23","2017/12/31")
	]
	print "\n".join(["%s\t%d" % q for q in computeworkdays(("2017/06/12",datetime.date.today()),holidays)])
if __name__ == '__main__':
	main()