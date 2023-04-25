import openpyexcel as px  # pip install openpyexcel
import re

wb = px.load_workbook('TT.xlsx')
sh = wb.active
start = (5, 1)
data = {'groups': []}
res_header = {}
weekdays = {'Понедельник': 'Monday',
            'Вторник': 'Tuesday',
            'Среда': 'Wednesday',
            'Четверг': 'Thursday',
            'Пятница': 'Friday',
            'Суббота': 'Saturday'
            }


def get_header_groups(row, col):
    val = ''
    for block_col in range(col, sh.max_column, 4):
        for tcol in range(block_col, block_col + 3):
            # if val != '' and val is not None:
            # print(sh.cell(row, tcol).value)
            tval = sh.cell(row, tcol).value
            if tval not in ('', None):
                val = sh.cell(row, tcol).value
                res_header[tcol] = val
                res_header[tcol + 1] = val
                res_header[tcol + 2] = val
                # print(sh.cell(row, tcol).coordinate)


def parse_block(row, col):
    for block_row in range(row, sh.max_row, 7):
        for start_block_col in range(col, sh.max_column, 4):
            pairs = []
            try:
                group = res_header[start_block_col + 1]
            except KeyError:
                break
            day = str(sh.cell(block_row, start_block_col).value).capitalize()
            if day != 'None':
                for row_pair in range(block_row, block_row + 6):
                    num = sh.cell(row_pair, start_block_col + 1).value
                    name = re.sub('  +', ' ', str(sh.cell(row_pair, start_block_col + 2).value))
                    name = re.sub('\\n.+', '', str(name))
                    # print(type(name))
                    if name != 'None':
                        cab = sh.cell(row_pair, start_block_col + 3).value
                        if cab is None:
                            cab = ''
                        pairs.append({'pairNum': num, 'pairName': name, 'pairCab': cab})
                # try:
                data['groups'].append({"groupName": group, "weekDay": weekdays[day], "pairs": pairs})
                # except KeyError:
                #     data['groups'][group] = {day: pairs}


row = 5
get_header_groups(row, 2)
parse_block(row + 2, 1)
print(str(data).replace("'", '"'))
