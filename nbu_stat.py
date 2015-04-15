import sys
import pdb
from mmap import mmap, ACCESS_READ
from pprint import pprint
from xlrd import open_workbook, xldate_as_tuple
from datetime import datetime

file_path = r'/home/skl1f/nbu_stat/Oper_stan_FEM.xls'


def collect_data(file_obj):
    values = []
    for s in file_obj.sheets():
        for row in range(s.nrows):
            dict1 = {'date': s.cell(row, 1).value,
                     'time': s.cell(row, 2).value,
                     'deals_count': s.cell(row, 3),
                     'amount': s.cell(row, 4),
                     'rate': s.cell(row, 5)}
            values.append(dict1)
    return values


def correct_date(data):
    placeholder = ''
    for x in range(len(data)):
        if isinstance(data[x]['date'], str) and placeholder == str(''):
            pass
        elif isinstance(data[x]['date'], str) and \
                isinstance(placeholder, tuple):
            data[x]['date'] = placeholder
        elif isinstance(data[x]['date'], float):
            del placeholder
            placeholder = xldate_as_tuple(data[x]['date'], 0)
            data[x]['date'] = placeholder
    return data


def correct_time(data):
    placeholder = (0, 0, 0, 17, 30, 0)
    for x in range(len(data)):
        if isinstance(data[x]['time'], str):
            data[x]['time'] = placeholder
        else:
            data[x]['time'] = xldate_as_tuple(data[x]['time'], 0)
    return data


def convert_data(data):
    placeholder = 0
    for x in range(len(data)):
        for i in ['amount', 'rate', 'deals_count']:
            if data[x][i].value == '-':
                data[x][i] = placeholder
            else:
                data[x][i] = data[x][i].value
    return data


def merge_datetime_correction(data):
    data = correct_date(data)
    data = correct_time(data)
    for x in range(len(data)):
        dt = list(data[x]['date'])
        dt[3:] = list(data[x]['time'])[3:]
        del data[x]['date']
        del data[x]['time']
        data[x]['datetime'] = datetime(*dt)
    return data


if __name__ == '__main__':
    if len(sys.argv) > 1:
        with open(sys.argv[1], 'rb') as f:

            f = open_workbook(
                file_contents=mmap(f.fileno(), 0,
                                   access=ACCESS_READ))

            data = collect_data(f)[7:]
            data = merge_datetime_correction(data)
            data = convert_data(data)
            pprint(data)
    else:
        print("Please ensure that path to xls-file are present.")
