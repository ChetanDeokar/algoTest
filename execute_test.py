import datetime
from rank_test import  WilcoxonRank

excel_obj = WilcoxonRank()
max_col = excel_obj.get_total_column()
max_row = excel_obj.get_total_row()
has_title = excel_obj.get_has_title()
first_column = excel_obj.get_column_data(1)

key_column = excel_obj.get_column_data(2)

key_column_data = None
key_column_header = None
first_column_data = None
first_column_header = None
final_data = {}
non_key_column_data = []
headers = []
if has_title:
    headers = excel_obj.get_row_data(1)

if isinstance(key_column, dict):
    key_column_data = key_column["data"]
    key_column_header = key_column["header"]

if isinstance(first_column, dict):
    first_column_data = first_column["data"]
    first_column_header = first_column["header"]


print "Started testing %s" % datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
if first_column_data and key_column_data:
    if max_col <= 2:
        print "Malformed data"
    else:
        range_end = max_row + 1
        if has_title:
            range_end = max_row - 1
        for col in xrange(3, max_col+1):
            column_data = None
            column_header = None
            current_column_data = excel_obj.get_column_data(col)
            if isinstance(current_column_data, dict):
                column_data = current_column_data["data"]
                column_header = current_column_data["header"]
            if column_data:
                difference = []
                non_key_column_data.append(column_data)
                for value in xrange(range_end):
                    difference.append((first_column_data[value], excel_obj.get_difference(key_column_data[value], column_data[value])))
                sorted_diff = sorted(difference, key=lambda x: abs(x[1]))
                #print "Getting ranked data"
                ranked_data = excel_obj.get_rank(sorted_diff)
                #print "Got ranked data"
                #print "Getting tied ranked data"
                tied_ranked_data = excel_obj.get_tied_rank(ranked_data)
                #print "Got tied ranked data"
                #print "Getting signed ranked data"
                signed_rank_data = excel_obj.get_signed_rank(tied_ranked_data)
                #print "Got signed ranked data"
                print "Started dumping data"
                for data in signed_rank_data:
                    identifier = data[0][0][0][0]
                    rank = data[1]
                    if identifier in final_data:
                        final_data[identifier].append(rank)
                    else:
                        final_data[identifier] = [rank]
                print "finished dumping data"
        if final_data:
            print "Writing Result"
            row_number = 1
            if headers:
                excel_obj.write_row(headers)
                row_number = 2
                print "Added headers"
            if first_column_data:
                excel_obj.write_column(first_column_data, column_number=1, row_number=row_number)
                print "Added first column"
            if key_column_data:
                excel_obj.write_column(key_column_data, column_number=2, row_number=row_number)
                print "Added key column"
            for index, temp_data in enumerate(non_key_column_data, start=3):
                excel_obj.write_column(temp_data, column_number=index, row_number=row_number)
            print "Added  non key column"
            rank_start_column_number = 3 + len(non_key_column_data)
            print "dumping ranks for %s rows in %s columns" % (len(final_data), len(non_key_column_data))
            for row_indx, data_id in enumerate(first_column_data):
                if data_id in final_data:
                    excel_obj.write_row(final_data[data_id], row_number=row_number+row_indx,
                                        column_number=rank_start_column_number)
            print "dumped ranks. Now saving to excel"
            excel_obj.save()
else:
    print "Error while reading excel or malformed excel"

print "Finished testing %s" % datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")