from read_excel import ReadExcel
from write_excel import WriteExcel


class Permutations(ReadExcel, WriteExcel):

    def __init__(self, excel_name="permutations.xlsx", has_title=True, result_file_name="permutation_result.xlsx"):
        if excel_name:
            ReadExcel.__init__(self, sheet_name=excel_name, has_title=has_title)
        else:
            ReadExcel.__init__(self, has_title=has_title)

        if result_file_name:
            WriteExcel.__init__(self, file_name=result_file_name)
        else:
            WriteExcel.__init__(self)

    def get_permutations(self, permutation_factor=1):
        total_columns = self.get_total_column()
        columns_for_permutations = total_columns - 2
        permutations = []
        headers = []
        start_index = 3
        for column in xrange(columns_for_permutations):
            column_data = self.get_column_data(start_index+column, self.get_has_title())
            column_data_list = None
            if isinstance(column_data, dict):
                column_data_list = column_data["data"]
            if column_data_list:
                next_column_index = start_index+column+1
                while next_column_index < total_columns+1:
                    next_column_data = self.get_column_data(next_column_index, self.get_has_title())
                    next_column_data_list = None
                    if isinstance(column_data, dict):
                        next_column_data_list = next_column_data["data"]
                    if next_column_data_list:
                        permutation_data = [(a*permutation_factor + b*permutation_factor)/2 for a, b in zip(column_data_list, next_column_data_list)]
                        if permutation_data:
                            permutations.append(permutation_data)
                            headers.append("W_Col%s_Col%s" % (start_index+column, next_column_index))
                    next_column_index += 1

        return permutations, headers
