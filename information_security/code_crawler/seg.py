# -*- coding: utf-8 -*-

import jieba
from collections import Counter
import pandas as pd
import excelSave as save
import os


def init_seg(file_name, cnt, data):
    for l in data['微博内容'].astype(str):
        seg_list = jieba.cut(l)  # 使用了jieba库进行分词，对分词结果进行简单统计
        for seg in seg_list:
            if seg not in STOPWORDS and seg not in PUNCTUATIONS:
                cnt[seg] = cnt[seg] + 1


# 将统计的词汇和数量写入excel表格
def save_seg(book_name_xls, sheet_name_xls, cnt, word_num):
    if os.path.exists(book_name_xls):
        print("文件已存在")
        value_title = [
            ["", "词语", "出现次数"], ]
        save.write_excel_xls(book_name_xls, sheet_name_xls, value_title)
    else:
        print("文件不存在，重新创建")
        value_title = [
            ["", "词语", "出现次数"], ]
        save.write_excel_xls(book_name_xls, sheet_name_xls, value_title)
    # f_out = open(book_name_xls, 'w+')
    result = cnt.most_common(word_num)  # 选出数量最多的前word_num个
    i = 0
    for ix in result:
        word = ix[0]  # 词语
        times = str(ix[1])  # 出现次数
        if len(word) >= 2:  # 只选择长度大于2的词汇存入excel
            value1 = [[i, word, times], ]
            save.write_excel_xls_append_canrepeat(book_name_xls, value1)
            i = i + 1


STOPWORDS = [u'的', u' ', u'\n', u'他', u'地', u'得', u'而', u'了', u'在', u'是', u'我', u'有', u'和', u'就', u'不', u'人', u'都',
             u'一', u'一个', u'上', u'也', u'很', u'到', u'说', u'要', u'去', u'你', u'会', u'着', u'没有', u'看', u'好', u'自己', u'这']
PUNCTUATIONS = [u'。', u'#', u'，', u'“', u'”', u'…', u'？', u'！', u'、', u'；', u'（', u'）']

if __name__ == '__main__':
    # 需要进行分词的文件
    file_name = '中考'
    cnt = Counter()
    data = pd.read_excel('weibodata/' + file_name + '.xls')
    init_seg(file_name, cnt, data)

    # 统计词汇出现次数
    word_num = 20  # 次数最多的词汇量
    book_name_xls = "seg_result/" + file_name + "话题微博词汇统计.xls"  # 填写你想存放excel的路径，没有文件会自动创建
    sheet_name_xls = '微博数据'  # sheet表名
    save_seg(book_name_xls, sheet_name_xls, cnt, word_num)
