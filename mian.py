# import csv
# with open("weather_data.csv") as datafile:
#     data = csv.reader(datafile)
#     for row in data:
#         print(row)


import pandas
data=pandas.read_csv("weather_data.csv")
print(data)
xls=pandas.ExcelWriter("Book1.xlsx")
data.to_excel(xls)
xls.save()