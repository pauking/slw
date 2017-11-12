#coding=utf-8
from __future__ import division
import pdb
import pandas as pd
from pandas import Series, DataFrame

# 开票基数
SUM_EACH_NOTE = 116990

# 票ID
note_id = 0


# 读取数据
def read_data(file_path):
    xls_file = pd.ExcelFile(file_path)
    table = xls_file.parse('Sheet1')
    return table


# 处理开票
def process(table):
    global note_id
    # 先进行排序 按找实际开票金额降序排列
    table = table.sort_values(by=u"实际开票金额", ascending=False)
    # 总共有几行
    num_rows = table.shape[0]
    # 开完每单后，当前剩余钱
    remaind_money = 0

    sum_money = 0


    for i in range(num_rows):
        # 结果字符串
        str_rlt = ""

        # 取出其中的某行
        # 列标签
        # 商品编码 商品名称 数量 单价 实际金额 票折 票折金额 票折单价 实际开票金额 备注（括号内为数量）

        i_row = table.iloc[i]

        # 判断当前 实际开票金额是否大于SUM_EACH_NOTE
        sjkpje = i_row[u'实际开票金额']

        # 单价
        price = i_row[u"票折单价"]

        if sjkpje > SUM_EACH_NOTE:

            #-----------处理第一张票------------

            cur_cn = (SUM_EACH_NOTE-remaind_money)//price
            str_rlt += str(int(note_id + 1)) + "(1*" + str(int(cur_cn)) + ")"

            # 更新note_id
            note_id += 1

            remaind_money = sjkpje - cur_cn * price
            # 剩余数
            remaind_cn = i_row[u"数量"] - cur_cn

            # 基于单价  SUM_EACH_NOTE 能最多开多少张单票
            most_dan_piao = SUM_EACH_NOTE // price

            # 开票数
            notes = remaind_cn // most_dan_piao
            if notes > 0:
                str_rlt +=u"、"+ str(int(note_id + 1)) + "-" + str(int(note_id + notes)) + "(" + str(int(notes)) + "*" + str(int(most_dan_piao)) + ")"
            # 更新票id
            note_id = note_id + notes

            # -------------处理最后一张票---------------
            # 剩余多少张开不出去
            remaind_dan_piao = remaind_cn - most_dan_piao * notes
            remaind_money = remaind_dan_piao * price
            if remaind_dan_piao > 0:
                str_rlt += u"、" + str(int(note_id + 1)) + "(1*" + str(int(remaind_dan_piao)) + ")"
                # note_id += 1
            print str_rlt


        else:
            remaind_money += sjkpje
            if remaind_money < SUM_EACH_NOTE:
                str_rlt = str(int(note_id+1)) + "(1*"+str(int(i_row[u"数量"]))+")"
                print str_rlt
            else:
                # 本次之前的剩余钱
                pre_money = remaind_money - sjkpje
                need_money = SUM_EACH_NOTE - pre_money

                need_cn = need_money // price
                str_rlt = str(int(note_id+1))+"(1*"+str(int(need_cn))+")"
                remaind_money = sjkpje - need_cn*price
                note_id += 1
                str_rlt += u"、" + str(int(note_id+1))+"(1*"+str(int(i_row[u"数量"]-need_cn))+")"
                print str_rlt

file_path = 'data/王.xlsx'
table = read_data(file_path)
process(table)
