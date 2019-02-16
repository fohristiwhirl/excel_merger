# Merges excel files, with some caveats:
#
#	- All files must have the same sheets
#	- Cells with data take priority over empty cells
#	- Otherwise, priority goes to earlier files in the args list

print("Loading openpyxl...")
import openpyxl, sys

def main():

	if len(sys.argv) < 2: # Allow a single file
		print("Usage: excel_merger.py <file1> <file2> <etc>")
		return

	workbooks = []

	for filename in sys.argv[1:]:
		print("Opening " + filename)
		workbooks.append(openpyxl.load_workbook(filename))

	if same_sheet_names(workbooks) == False:
		print("Workbooks had different sheets")
		return

	print("Merging...")
	merge(workbooks, "merged_output.xlsx")
	print("Done.")
	return


def same_sheet_names(workbooks):

	expected_sheets = workbooks[0].get_sheet_names()

	for workbook in workbooks[1:]:
		if workbook.get_sheet_names() != expected_sheets:
			return False

	return True


def merge(workbooks, outfilename):

	# Merge data into workbook 0 and then save it as a different filename
	# FIXME: can we create a copy of workbook 0 instead??

	sheet_names = workbooks[0].get_sheet_names()

	for name in sheet_names:

		target = workbooks[0].get_sheet_by_name(name)

		for workbook in workbooks[1:]:

			source = workbook.get_sheet_by_name(name)

			for y in range(1, source.max_row + 1):

				for x in range(1, source.max_column + 1):

					source_cell = source.cell(row = y, column = x)
					target_cell = target.cell(row = y, column = x)

					if target_cell.value is None and source_cell.value is not None:
						target_cell.value = source_cell.value

	workbooks[0].save(outfilename)


main()
