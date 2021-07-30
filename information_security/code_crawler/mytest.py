# -*- coding: utf-8 -*-
import unittest
from collections import Counter

import pandas as pd
import xlrd
from ddt import ddt, data, unpack
from selenium import webdriver

from analysis import isPostive
from create_weibo import water_army
from multipattern_matching import Trie
from normal_topic_spyder import spider
from seg import init_seg, save_seg
from visualization import Read_Excel, emotion_bar, emotion_pie, level_bar, level_pie, Read_Excel_keyword, keyword_bar, \
    keyword_pie


@ddt
class Cases(unittest.TestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.username = "15586430583"  # 你的微博登录名
        cls.password = "yutao19981119"  # 你的密码
        cls.maxWeibo = 20  # 设置最多多少条微博
        cls.sheet_name_xls = '微博数据'  # sheet表名

    def setUp(self) -> None:
        pass

    def tearDown(self) -> None:
        pass

    @data(['高兴', '积极'], ['恶心心', '消极'], ['今天是个好日子', '积极'],
          ['味道不错，确实不算太辣，适合不能吃辣的人。就在长江边上，抬头就能看到长江的风景。鸭肠、黄鳝都比较新鲜。', '积极'],
          ['内饰蛮年轻的 而且看上去质感都蛮好 貌似本田所有车都有点相似 满高档的', '积极'],
          ['竟然还敢说不喜欢奥迪 你的品味越来越令我失望 你怎么变成这个样子了', '消极'],
          ['院线看电影这么多年以来，这是我第一次看电影睡着了。简直是史上最大烂片！没有之一！侮辱智商！大家小心警惕！千万不要上当！再也不要看了！', '消极'])
    @unpack
    def test01_情感分析(self, text, emotion):
        self.assertEqual(emotion, isPostive(text))

    @data(['王思聪'])
    @unpack
    def test02_数据爬取(self, temp_filename):
        driver = webdriver.Chrome()  # 你的chromedriver的地址
        book_name_xls = "weibodata/" + temp_filename + ".xls"  # 填写你想存放excel的路径，没有文件会自动创建
        keywords = ["#" + temp_filename + "#"]  # 此处可以设置多个话题，#必须要加上
        for keyword in keywords:
            spider(self.username, self.password, driver, book_name_xls, self.sheet_name_xls, keyword,
                   self.maxWeibo)

    @data('王思聪')
    def test03_分词(self, temp_filename):
        cnt = Counter()
        data = pd.read_excel('weibodata/' + temp_filename + '.xls')
        init_seg(cnt, data)
        # 统计词汇出现次数
        word_num = 20  # 次数最多的词汇量
        book_name_xls = "seg_result/" + temp_filename + "话题微博词汇统计.xls"  # 填写你想存放excel的路径，没有文件会自动创建
        save_seg(book_name_xls, self.sheet_name_xls, cnt, word_num)

    @data('王思聪')
    def test04_情感分析2(self, temp_filename):
        file_path = "weibodata/" + temp_filename + ".xls"
        data = pd.read_excel(file_path)
        moods = []
        count = 1
        for i in data['微博内容']:
            moods.append(isPostive(i))
            count += 1
            print("目前分析到：" + str(count))
        data['情感倾向'] = pd.Series(moods)
        # 此处为覆盖保存
        data.to_excel(file_path)

    @data(['王思聪'])
    @unpack
    def test05_可视化(self, temp_filename):

        # 导入Excel 文件
        data = xlrd.open_workbook("weibodata/" + temp_filename + ".xls")
        # 载入第一个表格
        table = data.sheets()[0]
        tables = []
        Read_Excel(table, tables)
        emotion_bar(tables, temp_filename)
        emotion_pie(tables, temp_filename)
        level_bar(tables, temp_filename)
        level_pie(tables, temp_filename)
        # 词汇统计可视化
        temp_filename = temp_filename + "话题微博词汇统计"  # 文件名
        # 导入Excel 文件
        data = xlrd.open_workbook("seg_result/" + temp_filename + ".xls")
        # 载入第一个表格
        table = data.sheets()[0]
        tables = []
        Read_Excel_keyword(table, tables)
        keyword_bar(tables, temp_filename)
        keyword_pie(tables, temp_filename)

    @data(['王思聪', '5'])
    @unpack
    def test06_发微博(self, temp_filename, number):
        username = "16603620013"  # 你的微博登录名
        password = "wn123456789."  # 你的密码
        driver = webdriver.Chrome()  # 你的chromedriver的地址
        keywords = ["#" + temp_filename + "#"]  # 此处可以设置多个话题，#必须要加上
        weibo_num = int(number)  # 水军微博的数量
        # 载入第一个表格
        for keyword in keywords:
            for i in range(weibo_num):
                water_army(keyword, username, password, driver, i)

    @data([['中考', '延期'], '2020年中考：由于疫情中考原因延期了', [2, 1]],
          [['他', '她'], '他是她的他，她也是他的她', [3, 3]])
    @unpack
    def test07_多模式匹配(self, keywords, text, number):
        model = Trie(keywords)
        ret = model.search(text)
        count = 0
        temp = []
        for i in ret:
            temp.append(len(ret[i]))
            count = count + 1
        self.assertEqual(number, temp)


if __name__ == '__main__':
    unittest.main()
