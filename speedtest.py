import random, time
import openpyxl


def main():

	wb = openpyxl.Workbook()
	ws = wb.active
	populate(ws, 20, 5000)
	test(ws, cache = True)
	test(ws, cache = False)

	wb = openpyxl.Workbook()
	ws = wb.active
	populate(ws, 5000, 20)
	test(ws, cache = True)
	test(ws, cache = False)


def test(ws, cache = True):

	max_column = ws.max_column
	max_row = ws.max_row

	print("Data size is {} x {}, caching is {}".format(max_column, max_row, "ON" if cache else "OFF"))

	# Run the test one way...

	start_time = time.time()
	count = 0

	if cache:
		for x in range(1, max_column + 1):
			for y in range(1, max_row + 1):
				cell = ws.cell(column = x, row = y)
				count += 1
	else:
		for x in range(1, ws.max_column + 1):
			for y in range(1, ws.max_row + 1):
				cell = ws.cell(column = x, row = y)
				count += 1

	print("x-loop outer, time elapsed: {0:.2f} seconds ({1} accesses)".format(time.time() - start_time, count))

	# Run the test the other way...

	start_time = time.time()
	count = 0

	if cache:
		for y in range(1, max_row + 1):
			for x in range(1, max_column + 1):
				cell = ws.cell(column = x, row = y)
				count += 1
	else:
		for y in range(1, ws.max_row + 1):
			for x in range(1, ws.max_column + 1):
				cell = ws.cell(column = x, row = y)
				count += 1

	print("y-loop outer, time elapsed: {0:.2f} seconds ({1} accesses)".format(time.time() - start_time, count))


def populate(ws, width, height):
	for x in range(1, width + 1):
		for y in range(1, height + 1):
			if random.choice([True, False]):
				ws.cell(column = x, row = y).value = random.randint(0, 100)


main()
