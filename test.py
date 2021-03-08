import dateparser

# just a test
print("Hello World!")

d=dateparser.parse('2021-02-19T17:49:00.000Z')
print(str(d.year)," ",str(d.month)," ",str(d.day))