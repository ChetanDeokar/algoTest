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
