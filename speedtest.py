import random, time
import openpyxl

# Note: the speed problem turned out to be due to recalculation of either
# max_column or max_row; regardless of column or row size, each recalculation
# involves the library iterating over every cell; so ultimately the speed
# depended on the size of the inner loop. If the inner loop was run 5000 times
# then there would be ~5000 of these calculations.


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

	if cache:
		for x in range(1, max_column + 1):
			for y in range(1, max_row + 1):
				cell = ws.cell(column = x, row = y)
	else:
		for x in range(1, ws.max_column + 1):
			for y in range(1, ws.max_row + 1):
				cell = ws.cell(column = x, row = y)

	print("{0} cycles of TOP to BOTTOM, {1} accesses per cycle, time elapsed: {2:.2f} seconds".format(x, y, time.time() - start_time))

	# Run the test the other way...

	start_time = time.time()

	if cache:
		for y in range(1, max_row + 1):
			for x in range(1, max_column + 1):
				cell = ws.cell(column = x, row = y)
	else:
		for y in range(1, ws.max_row + 1):
			for x in range(1, ws.max_column + 1):
				cell = ws.cell(column = x, row = y)

	print("{0} cycles of LEFT to RIGHT, {1} accesses per cycle, time elapsed: {2:.2f} seconds".format(y, x, time.time() - start_time))


def populate(ws, width, height):
	for x in range(1, width + 1):
		for y in range(1, height + 1):
			if random.choice([True, False]):
				ws.cell(column = x, row = y).value = random.randint(0, 100)


main()
