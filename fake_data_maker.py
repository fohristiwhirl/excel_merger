import openpyxl, random

wb = openpyxl.Workbook()
ws = wb.active

for x in range(1, 20):
	for y in range(1, 5000):
		if random.choice([True, False]):
			ws.cell(column = x, row = y).value = random.randint(0, 100)

wb.save("fakedata.xlsx")
