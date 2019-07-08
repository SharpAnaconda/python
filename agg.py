from numpy import *
from openpyxl import load_workbook
from pandas import DataFrame, Series


def non_none_val(s, buf):
    if s is None:
        s = buf
    else:
        buf = s
    return s, buf


def create_column_titles(steps, elements, items):
    cols = []
    s_buf = ''
    el_buf = ''
    for s, el, i in zip(steps, elements, items):
        s, s_buf = non_none_val(s, s_buf)
        el, el_buf = non_none_val(el, el_buf)
        cols.append(str(i) + '_' + str(el) + '_' + str(s))
    return cols


def xl_to_df(file_path, p):
    data = load_workbook(filename=file_path, read_only=True).worksheets[p].values
    steps = next(data)
    elements = next(data)
    items = next(data)
    cols = create_column_titles(steps, elements, items)
    data = list(data)
    return DataFrame(data, columns=cols)


def group_all(df, col, o, app, ng, t, item, d):
    bf = DataFrame()
    bf[[t, item, d]] = df.groupby([col], as_index=False)[app, ng].sum().apply(
        lambda a: Series([col, a[col], zero_div_is_zero(a[ng], a[app])]), axis=1)
    return o.append(bf, sort=False)


def out_all(df, regex, app, ng, t, item, d):
    o = DataFrame()
    F = df.filter(regex=regex)
    for col in F:
        o = group_all(df, col, o, app, ng, t, item, d)
    return o


def create_sum_parts(df, f, part_use, rel_app, rel_ng, app, ng):
    c = 0
    for row in df.index:
        for column in f:
            b = df[column][row]
            for pc in f:
                if df[pc][row] == b:
                    part_num = pc.split("-")[1]
                    if df[part_use + "-" + part_num][row] is not None:
                        c += df[part_use + "-" + part_num][row]
            part_num = column.split("-")[1]
            if df[part_use + "-" + part_num][row] is None:
                df.loc[row, rel_app + "-" + part_num] = 0
                df.loc[row, rel_ng + "-" + part_num] = 0
            else:
                df.loc[row, rel_app + "-" + part_num] = df[app][row] * df[part_use + "-" + part_num][row] / c
                df.loc[row, rel_ng + "-" + part_num] = df[ng][row] * df[part_use + "-" + part_num][row] / c
            c = 0


def create_sum_parts_packs(f, df, part_name, rel_app, rel_ng):
    f = df.filter(regex='^' + part_name).filter(like='-')
    p = DataFrame()
    for column in f:
        part_num = column.split('-')[1]
        aggr = df.groupby([part_name + '-' + part_num], as_index=False)[
            rel_app + '-' + part_num, rel_ng + '-' + part_num].sum()
        af = aggr.filter(like='-')
        for a in af:
            aggr = aggr.rename(index=str, columns={a: a.split("-")[0]})
        p = p.append(aggr, sort=False)
    return p


def zero_div_is_zero(a, b):
    try:
        return a / b
    except ZeroDivisionError:
        return 0


def out_parts(t, item, d, p, part_name, o, rel_app, rel_ng):
    __bf = DataFrame()
    try:
        __bf[[t, item, d]] = p.groupby([part_name], as_index=False)[part_name, rel_app, rel_ng].sum().apply(
            lambda a: Series([part_name, a[part_name], zero_div_is_zero(a[rel_ng], a[rel_app])]), axis=1)
    except ValueError:
        pass
    return o.append(__bf, sort=False)


def main():
    FILENAME = 'Book1.xlsx'
    P = 0
    __df: DataFrame = xl_to_df(FILENAME, P)

    APP = '着手数量_実績_PE600'
    NG = '不良数量_実績_PE600'
    D = 'Degree'
    T = 'Type'
    ITEM = 'Item'
    __regex = '^(?!.*(' + APP + '|' + NG + ')).*$'
    __o = out_all(__df, __regex, APP, NG, T, ITEM, D)
    __o.to_csv('1.csv', encoding='utf-8', index=False)

    PART_PROPERTIES = ['部材種名', '部材機種コード', '部材ロットNo.', '外装部材ロットNo.', '内装部材ロットNo.']
    __part_name = PART_PROPERTIES[1]
    PART_USE = '使用数量'
    REL_APP = '該当_app'
    REL_NG = '該当_ng'

    __f = __df.filter(like=__part_name)
    create_sum_parts(__df, __f, PART_USE, REL_APP, REL_NG, APP, NG)

    __o = DataFrame()
    for __part_property in PART_PROPERTIES:
        __p = create_sum_parts_packs(__f, __df, __part_property, REL_APP, REL_NG)
        __o = out_parts(T, ITEM, D, __p, __part_property, __o, REL_APP, REL_NG)
    __o.to_csv('2.csv', encoding='utf-8', index=False)


main()
