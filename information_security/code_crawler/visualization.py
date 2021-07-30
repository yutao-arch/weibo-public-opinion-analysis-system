# coding=gbk
import xlrd
from pyecharts.charts import Bar
from pyecharts.charts import Pie
from pyecharts import options as opts


# 按行读取excel（爬取结果）的数据
def Read_Excel(table, tables):
    # 从第2行开始读取数据，因为这个Excel文件里面从第四行开始才是考生信息
    for rows in range(1, table.nrows - 1):
        dict_ = {"rid": "", "用户名称": "", "微博等级": "", "微博内容": "", "微博转发量": "", "微博评论量": "",
                 "微博点赞": "", "发布时间": "", "搜索关键词": "", "话题名称": "", "话题讨论数": "", "话题阅读数": "",
                 "情感倾向": ""}
        dict_["rid"] = table.cell_value(rows, 1)
        dict_["用户名称"] = table.cell_value(rows, 2)
        dict_["微博等级"] = table.cell_value(rows, 3)
        dict_["微博内容"] = table.cell_value(rows, 4)
        dict_["微博转发量"] = table.cell_value(rows, 5)
        dict_["微博评论量"] = table.cell_value(rows, 6)
        dict_["微博点赞"] = table.cell_value(rows, 7)
        dict_["发布时间"] = table.cell_value(rows, 8)
        dict_["搜索关键词"] = table.cell_value(rows, 9)
        dict_["话题名称"] = table.cell_value(rows, 10)
        dict_["话题讨论数"] = table.cell_value(rows, 11)
        dict_["话题阅读数"] = table.cell_value(rows, 12)
        dict_["情感倾向"] = table.cell_value(rows, 13)
        tables.append(dict_)


# 按行读取excel（词汇统计表）的数据
def Read_Excel_keyword(table, tables):
    # 从第2行开始读取数据，因为这个Excel文件里面从第四行开始才是考生信息
    for rows in range(1, table.nrows - 1):
        dict_ = {"词语": "", "出现次数": ""}
        dict_["词语"] = table.cell_value(rows, 1)
        dict_["出现次数"] = table.cell_value(rows, 2)
        tables.append(dict_)


# 用户情感可视化（柱状图）
def emotion_bar(tables, file_name):
    num_positive = 0
    num_negative = 0
    for i in tables:
        emotion = i["情感倾向"]
        if emotion == "积极":
            num_positive = num_positive + 1
        if emotion == "消极":
            num_negative = num_negative + 1
    bar_x_data = ("积极", "消极")
    bar_y_data = (num_positive, num_negative)
    c = (
        Bar()
            .add_xaxis(bar_x_data)
            .add_yaxis("微博数量", bar_y_data, color="#af00ff")
            .set_global_opts(title_opts=opts.TitleOpts(title=file_name + "话题情感分析柱状图", subtitle='制作人：余涛，王宁'))
            .render("chart_emotion/" + file_name + "话题情感分析柱状图.html")
    )
    print("情感分析柱状图绘制完成")


# 用户情感可视化（饼状图）
def emotion_pie(tables, file_name):
    num_positive = 0
    num_negative = 0
    for i in tables:
        emotion = i["情感倾向"]
        if emotion == "积极":
            num_positive = num_positive + 1
        if emotion == "消极":
            num_negative = num_negative + 1
    bar_x_data = ("积极", "消极")
    bar_y_data = (num_positive, num_negative)
    c = (
        Pie(init_opts=opts.InitOpts(height="800px", width="1200px"))
            .add("情感分析概览",
                 [list(z) for z in zip(bar_x_data, bar_y_data)],
                 center=["35%", "38%"],
                 radius="40%",
                 label_opts=opts.LabelOpts(
                     formatter="{b|{b}: }{c}  {per|{d}%}  ",
                     rich={
                         "b": {"fontSize": 16, "lineHeight": 33},
                         "per": {
                             "color": "#eee",
                             "backgroundColor": "#334455",
                             "padding": [2, 4],
                             "borderRadius": 2,
                         },
                     }
                 ))
            .set_global_opts(title_opts=opts.TitleOpts(title=file_name + "话题情感分析饼状图", subtitle='制作人：余涛，王宁'),
                             legend_opts=opts.LegendOpts(pos_left="0%", pos_top="65%"))
            .render("chart_emotion/" + file_name + "话题情感分析饼状图.html")
    )
    print("情感分析饼状图绘制完成")


# 用户类型可视化（柱状图）
def level_bar(tables, file_name):
    num_nomal_level = 0
    num_blue_v = 0
    num_yellow_v = 0
    num_gold_v = 0
    num_talent = 0
    for i in tables:
        level = i["微博等级"]
        if level == "普通用户":
            num_nomal_level = num_nomal_level + 1
        if level == "蓝v":
            num_blue_v = num_blue_v + 1
        if level == "黄v":
            num_yellow_v = num_yellow_v + 1
        if level == "金v":
            num_gold_v = num_gold_v + 1
        if level == "微博达人":
            num_talent = num_talent + 1
    bar_x_data = ("普通用户", "蓝v", "黄v", "金v", "微博达人")
    bar_y_data = (num_nomal_level, num_blue_v, num_yellow_v, num_gold_v, num_talent)
    c = (
        Bar()
            .add_xaxis(bar_x_data)
            .add_yaxis("微博数量", bar_y_data, color="#af00ff")
            .set_global_opts(title_opts=opts.TitleOpts(title=file_name + "话题用户等级柱状图", subtitle='制作人：余涛，王宁'))
            .render("chart_level/" + file_name + "话题用户等级柱状图.html")
    )
    print("用户等级柱状图绘制完成")


# 用户类型可视化（饼状图）
def level_pie(tables, file_name):
    num_nomal_level = 0
    num_blue_v = 0
    num_yellow_v = 0
    num_gold_v = 0
    num_talent = 0
    for i in tables:
        level = i["微博等级"]
        if level == "普通用户":
            num_nomal_level = num_nomal_level + 1
        if level == "蓝v":
            num_blue_v = num_blue_v + 1
        if level == "黄v":
            num_yellow_v = num_yellow_v + 1
        if level == "金v":
            num_gold_v = num_gold_v + 1
        if level == "微博达人":
            num_talent = num_talent + 1
    bar_x_data = ("普通用户", "蓝v", "黄v", "金v", "微博达人")
    bar_y_data = (num_nomal_level, num_blue_v, num_yellow_v, num_gold_v, num_talent)
    c = (
        Pie(init_opts=opts.InitOpts(height="800px", width="1200px"))
            .add("用户等级概览",
                 [list(z) for z in zip(bar_x_data, bar_y_data)],
                 center=["35%", "38%"],
                 radius="40%",
                 label_opts=opts.LabelOpts(
                     formatter="{b|{b}: }{c}  {per|{d}%}  ",
                     rich={
                         "b": {"fontSize": 16, "lineHeight": 33},
                         "per": {
                             "color": "#eee",
                             "backgroundColor": "#334455",
                             "padding": [2, 4],
                             "borderRadius": 2,
                         },
                     }
                 ))
            .set_global_opts(title_opts=opts.TitleOpts(title=file_name + "话题用户等级饼状图", subtitle='制作人：余涛，王宁'),
                             legend_opts=opts.LegendOpts(pos_left="0%", pos_top="65%"))
            .render("chart_level/" + file_name + "话题用户等级饼状图.html")
    )
    print("用户等级饼状图绘制完成")


# 词汇统计可视化（柱状图）
def keyword_bar(tables, file_name):
    keywords = []
    times = []
    for i in tables:
        keywords.append(i["词语"])
        times.append(i["出现次数"])
    bar_x_data = (keywords)
    bar_y_data = (times)
    c = (
        Bar()
            .add_xaxis(bar_x_data)
            .add_yaxis("词语出现次数", bar_y_data, color="#af00ff")
            .set_global_opts(title_opts=opts.TitleOpts(title=file_name + "柱状图", subtitle='制作人：余涛，王宁'))
            .render("chart_keyword/" + file_name + "柱状图.html")
    )
    print("词汇统计柱状图绘制完成")


# 词汇统计可视化（饼状图）
def keyword_pie(tables, file_name):
    keywords = []
    times = []
    for i in tables:
        keywords.append(i["词语"])
        times.append(i["出现次数"])
    bar_x_data = (keywords)
    bar_y_data = (times)
    c = (
        Pie(init_opts=opts.InitOpts(height="800px", width="1200px"))
            .add("词汇统计",
                 [list(z) for z in zip(bar_x_data, bar_y_data)],
                 center=["35%", "38%"],
                 radius="40%",
                 label_opts=opts.LabelOpts(
                     formatter="{b|{b}: }{c}  {per|{d}%}  ",
                     rich={
                         "b": {"fontSize": 16, "lineHeight": 33},
                         "per": {
                             "color": "#eee",
                             "backgroundColor": "#334455",
                             "padding": [2, 4],
                             "borderRadius": 2,
                         },
                     }
                 ))
            .set_global_opts(title_opts=opts.TitleOpts(title=file_name + "饼状图", subtitle='制作人：余涛，王宁'),
                             legend_opts=opts.LegendOpts(pos_left="0%", pos_top="65%"))
            .render("chart_keyword/" + file_name + "饼状图.html")
    )
    print("词汇统计饼状图绘制完成")


if __name__ == '__main__':
    # 情感分析和用户等级可视化
    file_name = "中考"  # 文件名
    # 导入Excel 文件
    data = xlrd.open_workbook("weibodata/" + file_name + ".xls")
    # 载入第一个表格
    table = data.sheets()[0]
    tables = []
    Read_Excel(table, tables)
    emotion_bar(tables, file_name)
    emotion_pie(tables, file_name)
    level_bar(tables, file_name)
    level_pie(tables, file_name)

    # 词汇统计可视化
    file_name = file_name + "话题微博词汇统计"  # 文件名
    # 导入Excel 文件
    data = xlrd.open_workbook("seg_result/" + file_name + ".xls")
    # 载入第一个表格
    table = data.sheets()[0]
    tables = []
    Read_Excel_keyword(table, tables)
    keyword_bar(tables, file_name)
    keyword_pie(tables, file_name)
