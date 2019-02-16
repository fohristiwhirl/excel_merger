import random, time
import openpyxl

print("Running test with tall rectangular data...")

wb = openpyxl.Workbook()
ws = wb.active

for x in range(1, 21):
	for y in range(1, 5001):
		if random.choice([True, False]):
			ws.cell(column = x, row = y).value = random.randint(0, 100)

# Run the test one way...

start_time = time.time()

for x in range(1, ws.max_column + 1):
	for y in range(1, ws.max_row + 1):
		cell = ws.cell(column = x, row = y)

print("x-loop outer, time elapsed: {0:.2f} seconds".format(time.time() - start_time))

# Run the test the other way...

start_time = time.time()

for y in range(1, ws.max_row + 1):
	for x in range(1, ws.max_column + 1):
		cell = ws.cell(column = x, row = y)

print("y-loop outer, time elapsed: {0:.2f} seconds".format(time.time() - start_time))

# -----------------------------------------------------------------------------------

print("Running test with wide rectangular data...")

wb = openpyxl.Workbook()
ws = wb.active

for x in range(1, 5001):
	for y in range(1, 21):
		if random.choice([True, False]):
			ws.cell(column = x, row = y).value = random.randint(0, 100)

# Run the test one way...

start_time = time.time()

for x in range(1, ws.max_column + 1):
	for y in range(1, ws.max_row + 1):
		cell = ws.cell(column = x, row = y)

print("x-loop outer, time elapsed: {0:.2f} seconds".format(time.time() - start_time))

# Run the test the other way...

start_time = time.time()

for y in range(1, ws.max_row + 1):
	for x in range(1, ws.max_column + 1):
		cell = ws.cell(column = x, row = y)

print("y-loop outer, time elapsed: {0:.2f} seconds".format(time.time() - start_time))
