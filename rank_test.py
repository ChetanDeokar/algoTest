from read_excel import ReadExcel


def get_difference(value1, value2):
    return value1 - value2

excel_obj = ReadExcel()
max_col = excel_obj.get_total_column()
max_row = excel_obj.get_total_row()
has_title = excel_obj.has_title()
first_column = excel_obj.get_column_data(1)

key_column = excel_obj.get_column_data(2)

key_column_data = None
key_column_header = None
first_column_data = None
first_column_header = None
final_data = {}
headers = []
if has_title:
    headers = excel_obj.get_row_data(1)

if isinstance(key_column, dict):
    key_column_data = key_column["data"]
    key_column_header = key_column["header"]

if isinstance(first_column, dict):
    first_column_data = first_column["data"]
    first_column_header = first_column["header"]

if first_column_data and key_column_data:
    if max_col <= 2:
        print "Malformed data"
    else:
        range_start = 1
        if has_title:
            range_start = 2
        for col in xrange(3, max_col+1):
            column_data = excel_obj.get_column_data(col)
            column_data = None
            column_header = None
            if isinstance(key_column, dict):
                column_data = key_column["data"]
                column_header = key_column["header"]
            if column_data:
                for value in xrange(range_start, max_row+1):
                    pass
else:
    print "Error while reading excel or malformed excel"

