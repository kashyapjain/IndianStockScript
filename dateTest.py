from datetime import date
import dateutil

i =[ date(2021,9,30), date(2021,10,1)]

j = 1
print(i[j].month > i[j-1].month)