import calendar

year = 2020
month = 7
day = 8

a = calendar.LocaleHTMLCalendar(locale='Russian_Russia')
with open('calendar.html', 'w') as g:
    print(a.formatyear(2014, width=4), file=g)


print (calendar.weekday(year, month, day))
print (calendar.monthrange(year, month))

for currDay in range (1, calendar.monthrange(year, month)[1]+1):
    print ( currDay,  calendar.weekday(year, month, currDay))
