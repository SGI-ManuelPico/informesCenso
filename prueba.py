import re
a = "2018-06-12 00:00:00"

print(re.findall('\d+',a.split("-")[2])[0])
