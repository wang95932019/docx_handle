#!/usr/bin/env python
# coding: utf-8
import json
import re

import pandas as pd
import xlrd3


def get_commit(df: pd.DataFrame) -> str:
    """ 获取工作表的备注 """
    return df.iloc[:, 0][df.iloc[:, 0].notnull()].values[-1]


def read_excel(path):
    """ 读取excel数据中的表1，表2，表3，表4 """

    ''' 读取存放课程基本信息的表1与表2 '''
    names1 = ['课程分类', '课程类别', '序号', '课程代码', '课程名称', '课程类型', '授课方式', '学分', '教学总学时',
              '实践学时', '1', '2', '3', '4', '5', '6', '说明']
    names2 = ['课程分类', '课程类别', '序号', '课程代码', '类别', '课程名称', '课程类型', '授课方式', '学分', '教学总学时',
              '实践学时', '1', '2', '3', '4', '5', '6', '说明']
    skiprows1, skiprows2 = 8, 4
    df1 = pd.read_excel(path, sheet_name=0, header=None, names=names1, skiprows=skiprows1, usecols=range(len(names1)))
    df2 = pd.read_excel(path, sheet_name=1, header=None, names=names2, skiprows=skiprows2, usecols=range(len(names2)))
    hours_sum = {}  # 创建字典保存总计数据

    # ''' 使用openpyxl的方式，处理合并单元格，但是openpyxl不能识别函数'''
    # import openpyxl
    #
    # def merge(sheet):
    #     # TODO 查询该sheet表单所有合并单元格
    #     merge_lists = sheet.merged_cells
    #     # print(merge_lists)
    #     merge_all_list = []
    #     # TODO 遍历合并单元格
    #     for merge_list in merge_lists:
    #         # TODO 获取单个合并单元格的起始行(row)和起始列(col)
    #         row_min, row_max, col_min, col_max = merge_list.min_row, merge_list.max_row, merge_list.min_col, merge_list.max_col
    #         if (row_min != row_max and col_min != col_max):
    #             row_col = [(x, y) for x in range(row_min, row_max + 1) for y in range(col_min, col_max + 1)]
    #             merge_all_list.append(row_col)
    #         elif (row_min == row_max and col_min != col_max):
    #             row_col = [(row_min, y) for y in range(col_min, col_max + 1)]
    #             merge_all_list.append(row_col)
    #         elif (row_min != row_max and col_min == col_max):
    #             row_col = [(x, col_min) for x in range(row_min, row_max + 1)]
    #             merge_all_list.append(row_col)
    #     return merge_all_list
    #     # TODO 得到的是个这样的列表值：[[(2, 1), (3, 1)], [(10, 1), (10, 2), (10, 3), (11, 1), (11, 2), (11, 3)]]
    #
    # wb = openpyxl.load_workbook(path)
    # sheet = wb.worksheets[1]
    # merge_lists = merge(sheet)
    # var1 = sheet['J5'].value
    # var2 = sheet.cell(row=5, column=10).value
    # for merge_list in merge_lists:
    #     min_row, min_col = merge_list[0]
    #     max_row, max_col = merge_list[-1]
    #     merge_len = len(merge_list)
    #     try:
    #         merge_value = float(sheet.cell(row=min_row, column=min_col).value)
    #         if min_col == max_col:
    #             cell_value = merge_value / merge_len
    #             df2.iloc[min_row-2:max_row+1-2, min_col - 1] = cell_value
    #     except ValueError:
    #         pass
    #     except TypeError:
    #         # except TypeError:TypeError: float() argument must be a string or a real number, not 'NoneType'
    #         pass
    # def ini_dfs(dfs: list, skiprowss: list):
    #
    #     """处理学分，教学总学时和各学期教学时分配的合并单元格
    #         和读取总计数据"""
    #
    #     workbook = xlrd3.open_workbook(path)
    #     for index, df in enumerate(dfs):
    #         hours_sum[f"{index}"] = {}  # 为每个工作表，在总计表中设置对应的索引
    #         # merged_cells 返回的是一个列表，每一个元素是合并单元格的位置信息的数组，数组包含四个元素（起始行索引，结束行索引，起始列索引，结束列索引）
    #         sheet = workbook.sheet_by_index(index)
    #         # 获取有合并单元格的实现方式
    #
    #         merge_cell_list = sheet.merged_cells  # （起始行索引，结束行索引，起始列索引，结束列索引）
    #         for (min_row, max_row, min_col, max_col) in merge_cell_list:
    #
    #             ''' 1. 处理学分，教学总学时和各学期教学时分配的合并单元格 '''
    #             # try:
    #             #     merge_value = float(sheet.cell_value(min_row, min_col))
    #             #
    #             #     if max_col - min_col == 1:
    #             #         cell_value = merge_value / (max_row - min_row)
    #             #         df.iloc[min_row - skiprowss[index]:max_row - skiprowss[index], min_col] = cell_value
    #             # except ValueError or TypeError:
    #             #     pass
    #             # except TypeError:
    #             #     # except TypeError:TypeError: float() argument must be a string or a real number, not 'NoneType'
    #             #     pass
    #
    #             ''' 2. 读取总计数据 '''
    #             merge_value = str(sheet.cell_value(min_row, min_col))
    #             if re.search('.*理论.*', merge_value):
    #                 if (min_row + 1, max_row + 1, min_col, max_col) in merge_cell_list:
    #                     try:
    #                         hours_sum[f"{index}"]["理论课"] = float(sheet.cell_value(min_row + 1, min_col))
    #                     except ValueError or TypeError:
    #                         pass
    #                     except TypeError:
    #                         # except TypeError:TypeError: float() argument must be a string or a real number, not 'NoneType'
    #                         pass
    #
    #             if re.search('.*实践.*', merge_value):
    #                 if (min_row + 1, max_row + 1, min_col, max_col) in merge_cell_list:
    #                     try:
    #                         hours_sum[f"{index}"]["实践课"] = float(sheet.cell_value(min_row + 1, min_col))
    #                     except ValueError or TypeError:
    #                         pass
    #                     except TypeError:
    #                         # except TypeError:TypeError: float() argument must be a string or a real number, not 'NoneType'
    #                         pass

    # ini_dfs([df1, df2], [8, 4])

    ''' 读取存放实践教学明细表的表3 '''
    names3 = ['学期', '课堂教学', '军事技能', '军事技能学时', '劳动教育', '课程实践',
              '认知实习', '岗位实习', '实习学时', '考试', '学期总周数',
              '注释1', '注释2', '注释3', '注释4', '注释5']

    df3 = pd.read_excel(path, sheet_name=2, header=None, names=names3, skiprows=4, usecols=range(len(names3)))

    ''' 读取存放课时学分统计表的表4 '''
    df4_1 = pd.read_excel(path, sheet_name=3, header=None, skiprows=4, nrows=6)
    df4_1.dropna(axis=1, inplace=True, how='all')
    df4_2 = pd.read_excel(path, sheet_name=3, header=None, skiprows=12, nrows=8)
    df4_2.dropna(axis=1, inplace=True, how='all')
    df4_2.dropna(axis=0, inplace=True, how='all')

    credit_statistics = {
        #     'course_type'
        'ggbxks': df4_1.iloc[0, 1],
        'ggbxxf': df4_1.iloc[1, 1],
        'ggbxxfbl': '{:.2%}'.format(df4_1.iloc[2, 1]),
        'zyjcks': df4_1.iloc[0, 2],
        'zyjcxf': df4_1.iloc[1, 2],
        'zyjcxfbl': '{:.2%}'.format(df4_1.iloc[2, 2]),
        'zyhxks': df4_1.iloc[0, 3],
        'zyhxxf': df4_1.iloc[1, 3],
        'zyhxxfbl': '{:.2%}'.format(df4_1.iloc[2, 3]),
        'sxks': df4_1.iloc[0, 4],
        'sxxf': df4_1.iloc[1, 4],
        'sxxfbl': '{:.2%}'.format(df4_1.iloc[2, 4]),
        'ggxxks': df4_1.iloc[0, 5],
        'ggxxxf': df4_1.iloc[1, 5],
        'ggxxxfbl': '{:.2%}'.format(df4_1.iloc[2, 5]),
        'ggrxks': df4_1.iloc[0, 6],
        'ggrxxf': df4_1.iloc[1, 6],
        'ggrxxfbl': '{:.2%}'.format(df4_1.iloc[2, 6]),
        'zyrxks': df4_1.iloc[0, 7],
        'zyrxxf': df4_1.iloc[1, 7],
        'zyrxxfbl': '{:.2%}'.format(df4_1.iloc[2, 7]),
        'hjks': df4_1.iloc[0, 9],
        'hjxf': df4_1.iloc[1, 9],
        'hjxfbl': '{:.2%}'.format(df4_1.iloc[2, 9]),
        'ggjcks': df4_1.iloc[3, 1],
        'ggkbl': '{:.2%}'.format(df4_1.iloc[3, 4]),
        'zyks': df4_1.iloc[3, 7],
        'zykbl': '{:.2%}'.format(df4_1.iloc[3, 9]),
        'zkss': df4_1.iloc[4, 3],
        'llkss': df4_1.iloc[4, 6],
        'sjkss': df4_1.iloc[4, 8],
        'llksbl': '{:.2%}'.format(df4_1.iloc[5, 3]),
        'sjksbl': '{:.2%}'.format(df4_1.iloc[5, 7]),
        #     'culture_program'
        'ggkxf': df4_2.iloc[0, 2],
        'ggkxfbl': '{:.2%}'.format(df4_2.iloc[0, 4]),
        'zykxf': df4_2.iloc[1, 2],
        'zykxfbl': '{:.2%}'.format(df4_2.iloc[1, 4]),
        'sjggsjxf': df4_2.iloc[2, 2],
        'sjzysjxf': df4_2.iloc[3, 2],
        'sjjxzxf': df4_2.iloc[2, 3],
        'sjggsjxfbl': '{:.2%}'.format(df4_2.iloc[2, 4]),
        'sjzysjxfbl': '{:.2%}'.format(df4_2.iloc[3, 4]),
        'sjjxxfzbl': '{:.2%}'.format(df4_2.iloc[2, 5]),
        'bxkxf': df4_2.iloc[4, 2],
        'bxkxfbl': '{:.2%}'.format(df4_2.iloc[4, 4]),
        'xxkxf': df4_2.iloc[5, 2],
        'xxkxfbl': '{:.2%}'.format(df4_2.iloc[5, 4]),
        'zxf': df4_2.iloc[6, 2]
    }

    # 获取备注
    commits = {}
    for name, df in zip(['平台课程教学进程表', '模块课程教学进程表', '实践教学'], [df1, df2, df3]):
        commits[name] = get_commit(df)

    return df1, df2, hours_sum, df3, credit_statistics, commits


def manage_df_1_2(df):
    """ 处理df1和df2数据 """

    '''1. 删除列 '''
    # drop_column = ['序号']
    drop_column = []
    __df = df.drop(drop_column, axis=1)
    ''' 2. 处理空数据 '''
    # 删除序号为空的数据
    __df.dropna(subset=["序号"], inplace=True)
    # 填充空值
    value = {'1': 0, '2': 0, '3': 0, '4': 0, '5': 0, '6': 0, '实践学时': 0, '说明': '', '课程代码': ''}
    __df.fillna(value=value, inplace=True)
    # 处理合并的单元格
    __df.fillna(method='ffill', inplace=True)

    __df.loc[__df["序号"].str.contains("小计", na=False), "课程代码"] = __df.loc[__df["序号"].str.contains("小计", na=False), "序号"]
    __df.loc[__df["序号"].str.contains("小计", na=False), ["序号", "课程名称", "课程类型", "授课方式"]] = ""
    __df.loc[(__df["课程代码"].str.contains("学时", na=False)) | (__df["课程代码"].str.contains("课时", na=False)), "学分"] = ""
    __df.loc[(__df["课程代码"].str.contains("学分", na=False)), ["教学总学时", "实践学时"]] = ""
    return __df


def manage_df3(df):
    """ 处理df3数据 """
    __df_1 = df.dropna(subset=['课堂教学', '学期总周数'])
    __df_2 = df.dropna(subset=['军事技能', '劳动教育'])
    __df = pd.concat([__df_1, __df_2])
    __df.drop_duplicates(inplace=True)
    # 填充空值
    __df.fillna(value="", inplace=True)
    __df['注释'] = __df['注释1'] + __df['注释2'] + __df['注释3'] + __df['注释4'] + __df['注释5']
    __df.drop(['注释1', '注释2', '注释3', '注释4', '注释5'], axis=1, inplace=True)
    return __df


def main(path):
    df1, df2, hours_sum, df3, credit_statistics, commits = read_excel(path)
    df_1and2 = pd.concat([manage_df_1_2(df1), manage_df_1_2(df2)])
    course = json.loads(df_1and2.to_json(orient='records', force_ascii=False))
    practice = json.loads(manage_df3(df3).to_json(orient='records', force_ascii=False))
    # dataFrame转Json
    return course, hours_sum, practice, credit_statistics, commits


if __name__ == '__main__':
    scyx = r"E:\Project\人才培养方案\docx_handle\test_excel\市场营销专业专业2022级教学计划安排表.xlsx"
    dsj = r"E:\Project\人才培养方案\docx_handle\test_excel\03.2022年度大数据技术专业教学计划安排表（2022年模板）.xls"

    main(scyx)
