# coding:utf-8
import math

import numpy
import numpy as np
import datetime
from datetime import date
import formulas
import pandas as pd
from formulas.functions import wrap_ufunc, is_number, raise_errors, flatten
from formulas.functions.info import iserror

from util import ret_row_column, number_to_char, check_value

FUNCTIONS = formulas.get_functions()

formulas = formulas

STATE_DATE = date(1900, 1, 1).toordinal() - 2


def xdata(y, m, d):
    # 1900-1-1  到 date 的天数值
    return date(y, m, d).toordinal() - STATE_DATE


def xmonth(date_number):
    # 返回当前数值天数的 月份
    date_str = str(date.fromordinal(date_number + STATE_DATE))
    return int(date_str.split("-")[1])


def xyeas(date_number):
    # 返回当前数值天数的 年份
    date_str = str(date.fromordinal(date_number + STATE_DATE))
    return int(date_str.split("-")[0])


def xedate(start_date, months):
    date_str = str(date.fromordinal(start_date + STATE_DATE))
    d = int(date_str.split("-")[2])
    new_date = xeomonth(start_date, months)
    date_str = str(date.fromordinal(new_date + STATE_DATE))
    new_y = int(date_str.split("-")[0])
    new_m = int(date_str.split("-")[1])
    new_d = int(date_str.split("-")[2])
    if d > new_d:
        d = new_d
    y = new_y
    m = new_m
    return xdata(y, m, d)


def xNOW():
    from datetime import datetime
    b = datetime.now()
    return datetime.date(b).toordinal() - STATE_DATE


def xeomonth(start_date, months):
    start_date = str(date.fromordinal(start_date + STATE_DATE))
    y = int(start_date.split("-")[0])
    m = int(start_date.split("-")[1])
    d_31 = [1, 3, 5, 7, 8, 10, 12]
    lists = [i for i in range(1, 13)]
    m = m + int(months)
    if m <= 0:
        m -= 1
        m = lists[m]
        y -= 1
    elif m > 12:
        m, mor_y = month_to_y(m)
        y += mor_y
    day_list = {}
    for i in range(1, 13):
        if i in d_31:
            day_list[i] = 31
        else:
            day_list[i] = 30
    if ((y % 4 == 0 and y % 100 != 0) or (y % 400 == 0 and y % 3200 != 0)):
        day_list[2] = 29
    else:
        day_list[2] = 28
    d = day_list[m]

    return xdata(y, m, d)


def month_to_y(m, y=0):
    if m < 12:
        return m, y
    m = m - 12
    y += 1
    return month_to_y(m, y)


def xoffset(reference, rows=0, cols=0, height=0, width=0):
    """
    =Offset（参照单元格，行偏移量，列偏移量，返回几行，返回几列）
    :return: 引用("=A1")
    """

    rows = 0 if iserror(rows) else rows
    cols = 0 if iserror(cols) else cols
    height = 0 if iserror(height) else height
    width = 0 if iserror(width) else width
    height -= 1
    width -= 1
    if height < 0:
        height = 0
    if width < 0:
        width = 0
    if "!" in reference:
        sp = reference.split("!")
    else:
        sp = ["", reference]
    row, column = ret_row_column(sp[1])
    row += int(rows)
    column += int(cols)
    char = number_to_char(column)
    # 偏移后位置
    # 添加几行几列
    row, column = ret_row_column(char + str(row))
    row_list = [row, row + int(height)]
    column_list = [column, column + int(width)]
    row_list.sort()
    column_list.sort()
    if sp[0]:
        sheet = sp[0] + "!"
    else:
        sheet = ''
    ret_str = "=" + sheet + number_to_char(column_list[0]) + str(row_list[0]) + ":" \
              + number_to_char(column_list[1]) + str(row_list[1])
    return ret_str


def xchar(date: int):
    return chr(date)


def xINDIRECT(data):
    pass


def xvalue(text):
    # print(text)
    return float(text)


def xVLOOKUP(lookup_value, table_array, col_index_num, range_lookup=0):
    col_index_num = int(col_index_num) - 1
    df = pd.DataFrame(table_array)
    try:
        val = df.loc[df[0] == lookup_value][col_index_num]
        val = val.tolist()[0]
    except:
        try:
            val = df.loc[df[0] == str(lookup_value)][col_index_num]
            val = val.tolist()[0]
        except Exception as e:
            print("查询出现问题")
            print(e)
            val = 0
    try:
        val = float(val)
    except:
        return val
    return val


def xsumif(range, criteria, sum_range):
    # import re
    # criteria 可以存着大于小于号等逻辑运算符。需要判断
    # regex = re.compile("^\W+")
    # res = regex.findall(criteria)
    # dict_s = {
    #     '<': lambda x, y: x < y,
    #     '<=': lambda x, y: x <= y,
    #     '>': lambda x, y: x > y,
    #     '>=': lambda x, y: x >= y,
    #     '=': lambda x, y: x == y,
    #     '<>': lambda x, y: x != y
    # }
    func = lambda x, y: x == y
    # if res:
    # func = dict_s.setdefault(res[0], lambda x, y: x == y)
    # regex2 = re.compile("\w+$")
    # criteria = ">=23"
    # criteria = eval(regex2.findall(criteria)[0])
    sums = 0
    for i, v in enumerate(range):
        if func(v, criteria):
            if sum_range[i]:
                sums += float(sum_range[i])
    return sums


def xPPMT(rate=0, per=0, nper=0, pv=0, fv=0, type1=0):
    # 一直用等额本金  默认0
    per = check_value(type(per))(per)
    if per == 0:
        return 0
    rate = check_value(type(rate))(rate)
    nper = check_value(type(nper))(nper)
    pv = check_value(type(pv))(pv)
    fv = check_value(type(fv))(fv)
    try:
        a = numpy.ppmt(rate, per, nper, pv, fv)
    except:
        a = 0
    return a


def xsum(*args):
    raise_errors(args)
    for i in args:
        sum(float(i))

    return sum(list(flatten(args)))


def countif(x, y):
    return 1


def xtranspose(table_array):
    vec = np.matrix(table_array)
    vec = vec.T
    return vec


def xiferror2222(val, error_value=0):
    return error_value if iserror(val) else val


def xround(x, d=0, func=round):
    d = 10 ** int(d)
    v = func(abs(x * d)) / d
    return -v if x < 0 else v


def xIPMT(a=0, b=0, c=0, d=0):
    return 0


def xnpv2(rate, cashflows):
    return sum([cf / (1 + rate) ** ((t - cashflows[0][0]).days / 365.0) for (t, cf) in cashflows])


def xnpv(rate, value, dates):
    cashflows = []
    try:
        for i, v in enumerate(dates):
            cashflows.append((datetime.datetime.fromordinal(v + STATE_DATE), value[i - 1]))
        return sum([cf / (1 + rate) ** ((t - cashflows[0][0]).days / 365.0) for (t, cf) in cashflows])
    except Exception as e:
        # print("xnpv error")
        # print(e)
        return 0


def secant_method(tol, f, x0):
    x1 = x0 * 1.1
    while (abs(x1 - x0) / abs(x1) > tol):
        x0, x1 = x1, x1 - f(x1) * (x1 - x0) / (f(x1) - f(x0))
    return x1


def xnpv_qitayong(rate, cashflows):
    try:
        chron_order = sorted(cashflows, key=lambda x: x[0])
        t0 = chron_order[0][0]  # t0 is the date of the first cash flow
        result = sum([cf / (1 + rate) ** ((t - t0).days / 365.0) for (t, cf) in chron_order])
        return result
    except Exception as e:
        # print(e)
        return 0


def xirr(value, dates, guess=0.1):
    cashflows = []
    try:
        for i, v in enumerate(dates):
            datass = 0
            if isinstance(value[i - 1], np.float64):
                datass = float(value[i - 1])
            if isinstance(value[i - 1], np.int32):
                datass = int(value[i - 1])
            x, _ = math.modf(datass)
            if x == 0:
                datass = int(value[i - 1])
            cashflows.append((datetime.datetime.fromordinal(v + STATE_DATE), datass))
        return secant_method(0.0001, lambda r: xnpv_qitayong(r, cashflows), guess)
    except Exception as e:
        # print(e)
        return 0


def xROUNDUP(x, d=0):
    d = 10 ** int(d)
    v = math.ceil(abs(x * d)) / d
    return -v if x < 0 else v


def xsumproduct(*args):
    # Check all arrays are the same length
    # Excel returns #VAlUE! error if they don't match
    raise_errors(args)
    for i in args:
        for j in i:
            if j[0] == "":
                j[0] = 0
    assert len(set(arg.size for arg in args)) == 1
    inputs = []
    for a in args:
        a = a.ravel()
        x = np.zeros_like(a, float)
        b = np.vectorize(is_number)(a)
        x[b] = a[b]
        inputs.append(x)

    return np.sum(np.prod(inputs, axis=0))


FUNCTIONS['DATE'] = xdata
FUNCTIONS['YEAR'] = xyeas
FUNCTIONS['OFFSET'] = xoffset
FUNCTIONS['EOMONTH'] = xeomonth
FUNCTIONS['EDATE'] = xedate
FUNCTIONS['CHAR'] = xchar
FUNCTIONS['INDIRECT'] = xINDIRECT
FUNCTIONS['VALUE'] = xvalue

FUNCTIONS['COUNTIF'] = countif
FUNCTIONS['TRANSPOSE'] = xtranspose
FUNCTIONS['IFERROR'] = {
    'function': wrap_ufunc(
        xiferror2222, input_parser=lambda *a: a
    )}

FUNCTIONS['ROUND'] = wrap_ufunc(xround)

# 为啥启用，数据必须都是str才能比较
FUNCTIONS['VLOOKUP'] = xVLOOKUP
FUNCTIONS['SUMIF'] = xsumif
FUNCTIONS['MONTH'] = xmonth
FUNCTIONS['PPMT'] = xPPMT
FUNCTIONS['IPMT'] = xIPMT
FUNCTIONS['ROUNDUP'] = xROUNDUP
FUNCTIONS['NOW'] = xNOW
FUNCTIONS['XIRR'] = xirr
FUNCTIONS['XNPV'] = xnpv

# 为啥注释因为 输出参数为array
# FUNCTIONS['SUMPRODUCT'] = wrap_func(xsumproduct)


if __name__ == '__main__':
    pass
    # xtranspose([[1, 2, 3]])
    a = [-2149.0224, -2149.0224, -2149.0224, -2149.0224, -2149.0224, -2149.0224, -2149.0224, -2149.0224, -2149.0224,
         -2149.0224, -2149.0224, -2259.4224, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739,
         320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739,
         320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739,
         320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 320.4739, 297.4575, 297.4575, 297.4575, 297.4575,
         297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 297.4575,
         297.4575, 297.4575, 297.4575, 297.4575, 297.4575, 272.2752, 270.6372, 270.6372, 270.6372, 261.8944, 261.8944,
         261.8944, 261.8944, 261.8944, 261.8944, 261.8944, 261.8944, 261.8944, 261.8944, 261.8944, 261.8944, 238.0638,
         238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0639, 238.0639, 238.0639,
         238.0639, 238.0639, 238.0639, 238.0639, 238.0639, 238.0639, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638,
         238.0638, 238.0638, 238.0638, 238.0638, 238.0639, 238.0639, 238.0639, 238.0639, 238.0639, 238.0639, 238.0639,
         238.0639, 238.0639, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638, 238.0638,
         238.0639, 238.0639, 238.0639, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57,
         230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57,
         230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57,
         230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57,
         230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 230.57, 223.0762, 223.0762, 223.0762, 223.0762,
         223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762,
         223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762,
         223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762,
         223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762,
         223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762, 223.0762,
         1494.2715, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    # data = [(datetime.date(2006, 1, 24), -39967), (datetime.date(2008, 2, 6), -19866),
    #         (datetime.date(2010, 10, 18), 245706), (datetime.date(2013, 9, 14), 52142)]
    # print(xirr(data))
    print(xPPMT(0.1 / 12, 1, 2 * 12, 2000))
