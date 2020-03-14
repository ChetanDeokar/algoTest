from read_excel import ReadExcel
from write_excel import WriteExcel


class WilcoxonRank(ReadExcel, WriteExcel):

    def __init__(self, excel_name=None, has_title=True, result_file_name=None):
        if excel_name:
            ReadExcel.__init__(self, sheet_name=excel_name, has_title=has_title)
        else:
            ReadExcel.__init__(self, has_title=has_title)

        if result_file_name:
            WriteExcel.__init__(self, file_name=result_file_name)
        else:
            WriteExcel.__init__(self)

    def get_difference(self, value1, value2):
        return value1 - value2

    def get_rank(self, sorted_list):
        rank = []
        rnk = 1
        for element in sorted_list:
            if abs(element[1]) > 0:
                rank.append((element, rnk))
                rnk += 1
            else:
                rank.append((element, 0))
        rank = sorted(rank, key=lambda x: x[1])
        return rank

    def get_tied_rank(self, rank_list):
        result = []
        jump_to_index = None
        for index, element in enumerate(rank_list):
            if jump_to_index and index < jump_to_index:
                continue
            rank = element[1]
            diff = abs(element[0][1])
            if rank > 0:
                to_be_processed_index = [index]
                to_be_averaged_rank = [rank]
                if index == (len(rank_list) - 1):
                    result.append((element, rank))
                else:
                    for indx, elm in enumerate(rank_list[index+1:], start=1):
                        current_elm_diff = abs(elm[0][1])
                        if diff == current_elm_diff:
                            to_be_processed_index.append(indx+index)
                            to_be_averaged_rank.append(elm[1])
                        else:
                            break
                    if to_be_processed_index and to_be_averaged_rank:
                        new_rank = sum(to_be_averaged_rank) / float(len(to_be_averaged_rank))
                        for index_to_update in to_be_processed_index:
                             result.append((rank_list[index_to_update], new_rank))

                        jump_to_index = (max(to_be_processed_index) + 1)
            else:
                result.append((element, 0))
        result = sorted(result, key=lambda x: x[1])
        return result

    def get_signed_rank(self, tied_ranked_list):
        rank = []
        for element in tied_ranked_list:
            current_rank = element[1]
            diff = element[0][0][1]
            if diff > 0:
                rank.append((element, current_rank*1))
            elif diff < 0:
                rank.append((element, current_rank * -1))
            else:
                rank.append((element, 0))
        rank = sorted(rank, key=lambda x: x[1])
        return rank


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
                for value in xrange(range_end):
                    difference.append((first_column_data[value], excel_obj.get_difference(key_column_data[value], column_data[value])))
                sorted_diff = sorted(difference, key=lambda x: abs(x[1]))
                ranked_data = excel_obj.get_rank(sorted_diff)
                tied_ranked_data = excel_obj.get_tied_rank(ranked_data)
                signed_rank_data = excel_obj.get_signed_rank(tied_ranked_data)
                for data in signed_rank_data:
                    identifier = data[0][0][0][0]
                    rank = data[1]
                    if identifier in final_data:
                        final_data[identifier].append(rank)
                    else:
                        final_data[identifier] = [rank]
        if final_data:
            row_number = 1
            if headers:
                excel_obj.write_row(headers)
                row_number = 2
            if first_column_data:
                excel_obj.write_column(first_column_data, column_number=1, row_number=row_number)
            if key_column_data:
                excel_obj.write_column(key_column_data, column_number=2, row_number=row_number)
            for data_id in first_column_data:
                if data_id in final_data:
                    excel_obj.write_row(final_data[data_id], row_number=row_number, column_number=3)
            excel_obj.save()
else:
    print "Error while reading excel or malformed excel"

