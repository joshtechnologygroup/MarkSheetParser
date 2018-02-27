from xlrd import open_workbook
import itertools
import traceback

available_formats = []
path = '/' # Path to the excel file containing students data


def calculate_source_combinations(horizontal_sources=None, vertical_sources=None):
    """Function to calculate all the possible sources combinations"""
    source_cf_dict = {}
    block_delimiter = '#'

    # 1. Compute source combinations for horizontal sources
    for i in range(1, 2):
        ith_combinations = itertools.combinations(horizontal_sources, i)
        for comb in ith_combinations:
            key = block_delimiter.join(list(comb))
            source_cf_dict[key] = 0.0

    # 2. Compute source combinations for vertical sources
    for i in range(1, 2):
        ith_combinations = itertools.combinations(vertical_sources, i)
        for comb in ith_combinations:
            key = block_delimiter.join(list(comb))
            source_cf_dict[key] = 0.0
    return source_cf_dict


def check_in_available_formats(horizontal_sources=None, vertical_sources=None):
    """check if the given sources list is present in the available formats"""
    formats_dict = calculate_source_combinations(horizontal_sources, vertical_sources)
    current_format_keys_set = set(formats_dict.keys())
    found = False

    for available_format_dict in available_formats:
        available_format_keys_set = set(available_format_dict.keys())

        if len(current_format_keys_set - available_format_keys_set) == 0:
            found = True
    if not found:
        available_formats.append(formats_dict)
    return found, formats_dict


def compute_sum_between_cells(start_row, start_col, end_row, end_col, sheet):
    count = 0.0
    sum = 0.0
    if start_col == end_col:
        for index in range(start_row, end_row + 1):
            try:
                value = float(sheet.cell(index, start_col).value)
                sum = sum + value
                count = count + 1
            except: pass

    elif start_row == end_row:
        for index in range(start_col, end_col + 1):
            try:
                value = float(sheet.cell(start_row, index).value)
                sum = sum + value
                count = count + 1
            except: pass
    return sum, count


def compute_source_average(key, sheet):
    sources = key.split('#')
    total_sum = 0
    total_count = 0
    for source in sources:
        start_point, end_point = source.split('_')
        start_row, start_col = start_point.split('*')
        end_row, end_col = end_point.split('*')
        source_sum, source_count = compute_sum_between_cells(
            int(start_row), int(start_col), int(end_row), int(end_col), sheet)
        total_sum = total_sum + source_sum
        total_count = total_count + source_count
    return total_sum / (total_count or 1)


def compute_confidence(sources_dict, sheet, aggregate):
    for key in sources_dict.keys():
        try:
            avg = compute_source_average(key, sheet)
        except Exception as err:
            print err.message, traceback.print_exc()
        error = abs(avg - aggregate)
        score = (10 / (error or 1))

        if error > 2:
            score = 0.1
        sources_dict[key] = sources_dict[key] + score


def compute_horizontal_souces(s, start_row_no, start_col_no, aggregate, visited_rows, h_sources):
    start_point = start_col_no
    for col in range(start_point, s.ncols):
        try :
            value = float(s.cell(start_row_no, col).value)
            if not start_point:
                start_point = col
        except :
            if start_point:
                h_sources.append(
                    '{}*{}_{}*{}'.format(start_row_no, start_point, start_row_no, col)
                )
            start_point = None


def compute_vertical_sources(s, start_row_no, start_col_no, aggregate, visited_rows, v_sources):
    start_point = start_col_no
    for row in range(start_point, s.nrows):
        try :
            value = float(s.cell(row, start_col_no).value)
            if not start_point:
                start_point = row
        except :
            if start_point:
                v_sources.append(
                    '{}*{}_{}*{}'.format(start_point, start_col_no, row, start_col_no)
                )
            start_point = None


def compute_sources(s, start_row_no, start_col_no, aggregate, visited_rows, visited_columns, h_sources, v_sources):
    if (start_row_no in visited_rows) or (start_col_no in visited_columns):
        return
    visited_rows.append(start_row_no)
    visited_columns.append(start_col_no)
    #compute_horizontal_souces(s, start_row_no, start_col_no, aggregate, visited_rows, h_sources)
    compute_vertical_sources(s, start_row_no, start_col_no, aggregate, visited_rows, v_sources)


def train(aggregate, marksheet_name):
    wb = open_workbook(path.format(marksheet_name + '.xlsx'))
    s = wb.sheets()[0]

    visited_rows = []
    visited_columns = []
    h_sources = []
    v_sources = []

    for row in range(s.nrows):
        for col in range(s.ncols):
            try:
                float(s.cell(row, col).value)
                compute_sources(s, row, col, aggregate, visited_rows, visited_columns, h_sources, v_sources)
                _, forma = check_in_available_formats(h_sources, v_sources)
                compute_confidence(forma, s, aggregate)
            except:
                pass


if __name__ == "__main__":
    path = '/home/ubuntu/Desktop/marksheets/10th/training_marksheets/{}'
    wb = open_workbook(path.format('students_data.xlsx'))
    s = wb.sheets()[0]
    for row in range(1, s.nrows):
        train(s.cell(row, 4).value, s.cell(row, 3).value)
