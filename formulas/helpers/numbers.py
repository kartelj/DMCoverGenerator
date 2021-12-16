# -*- coding: utf-8 -*- 

def calculate_percente(a1,total):
    try:
        value = (a1/total)*100
    except:
        value = 0
    return "{0} %".format(round(value, 2))

def calculate_percente_round_1(a1, total):
    try:
        value = (a1/total)*100
    except:
        value = 0
    return "{0} %".format(round(value, 1))

def calculate_percente_2(a1, total):
    try:
        value = (a1/total)*100
    except:
        value = 0
    return "{0} %".format(round(value))

# def convert_to_percente(a1, total):
#     try:
#         value = a1/total
#     except:
#         value = 0
#     return "{0} %".format(round(value))

def calculate_percente_3(a1,a2,a3):
    try:
        value = a1/a2*a3
    except:
        value = 0
    return "{0} %".format(round(value))

def div(a1, total):
    try:
        value = a1/total
    except:
        value = 0
    return round(value, 2)

def div_2(a1, total):
    try:
        value = a1/total*100
    except:
        value = 0
    return value

def div_3(a1,total):
    try:
        value = a1/total*100
    except:
        value = 0
    return round(value)


def div_round_1(a1,total):
    try:
        value = a1/total*100
    except:
        value = 0
    return value

def float_with_comma(num):
    # return "{:,}".format(num)
    return num

# def float_to_percente(num):
#     res = round(num*100,2)
#     return "{0} %".format(res)

def percente_to_float(num):
    return float(num[:-1])

def sum_values_for_quarter(value, quarter):
    if quarter == 1:
        return value[0]+value[1]+value[2]
    if quarter == 2:
        return value[3]+value[4]+value[5]
    if quarter == 3:
        return value[6]+value[7]+value[8]
    if quarter == 4:
        return value[9]+value[10]+value[11]


def calculate_days(mysum):
    return mysum.days + round(mysum.seconds / 86400)

def calculate_days_float(mysum):
    return mysum.days + round(mysum.seconds / 86400, 2)


def get_quarter_header(year,quarter,with_months=False,current_quarter=True):
    if with_months:
        return quarter_header_with_months(year,quarter,current_quarter)
    else:
        return regular_quarter_header(year,quarter)

def regular_quarter_header(year,quarter):
    h = ["{0} FY".format(year-1)]
    if quarter == 1:
        h += ["{0} Q1".format(year)]
    elif quarter == 2:
            h += ["{0} Q1".format(year),"{0} Q2".format(year)]
    elif quarter == 3:
        h += ["{0} Q1".format(year), "{0} Q2".format(year), "{0} Q3".format(year)]
    elif quarter == 4:
        h += ["{0} Q1".format(year), "{0} Q2".format(year), "{0} Q3".format(year),"{0} Q4".format(year) ]
    h += ["{0} ytd".format(year)]
    return h

def quarter_header_with_months(year,quarter,current_quarter):
    h = ["{0} FY".format(year-1)]
    if quarter == 1:
        if current_quarter:
            h += ["JAN","FEB","MAR","{0} Q1".format(year)]
        else:
            h += ["JAN","FEB","MAR"]
    elif quarter == 2:
        if current_quarter:
            h += ["{0} Q1".format(year),"APR","MAJ","JUN","{0} Q2".format(year)]
        else:
            h += ["{0} Q1".format(year),"APR","MAJ","JUN"]
    elif quarter == 3:
        if current_quarter:
            h += ["{0} Q1".format(year), "{0} Q2".format(year),"JUL","AVG","SEP","{0} Q3".format(year)]
        else:
            h += ["{0} Q1".format(year), "{0} Q2".format(year),"JUL","AVG","SEP"]
    elif quarter == 4:
        if current_quarter:
            h += ["{0} Q1".format(year), "{0} Q2".format(year), "{0} Q3".format(year),"OKT","NOV","DEC","{0} Q4".format(year)]
        else:
            h += ["{0} Q1".format(year), "{0} Q2".format(year), "{0} Q3".format(year),"OKT","NOV","DEC"]
    h += ["{0} ytd".format(year)]
    return h
