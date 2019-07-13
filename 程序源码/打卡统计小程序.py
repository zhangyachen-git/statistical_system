#!/usr/bin/env python
# coding: utf-8

# # 打卡数据分析

# 1.每个小组的应打卡人数
# 2.每个小组的实际打卡人数
# 3.每个小组的打卡率
# 4.小组的日积分  成员积分相加
# 5.小组的日均分  总积分/小组人员个数
# 6.统计日积分的前三，学号，姓名，分数
# 7.当日的小组平均分排名前三名

import os
import warnings
# ## 引入工具包
from datetime import date, datetime, timedelta
from decimal import Decimal

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')


def sratistics():
    # 定义输出文件夹
    output_path = 'output/'
    ## 读取文件
    today = date.today().strftime("%Y%m%d")
    print('今天的日期：'+today)
    file_name = 'data/'+today+'.xls'
    print('文件路径/文件名称：'+file_name)
    # file_name = "data/20190712.xlsx"
    if os.path.isfile(file_name):
        data = pd.read_excel(file_name)
        # ## 时间筛选
        yesterday = (date.today() + timedelta(days=-1)
                     ).strftime("%Y-%m-%d")    # 昨天日期
        print('数据的日期：'+yesterday)
        submitTime1 = pd.to_datetime(
            yesterday+' 00:00:00', format='%Y-%m-%d %H:%M:%S')
        submitTime2 = pd.to_datetime(
            yesterday+' 23:59:59', format='%Y-%m-%d %H:%M:%S')
        # print(submitTime1, submitTime2)
        data2 = data[
            (pd.to_datetime(data['修改时间'], format='%Y-%m-%d %H:%M:%S') >= submitTime1) &
            (pd.to_datetime(data['修改时间'], format='%Y-%m-%d %H:%M:%S') <= submitTime2)]
        data3 = data2.drop(columns=[
            '参加的吸引力法则线下公开课详细信息', '参加的《与师父有约》落地读书会信息', '邀请了多少位新朋友参加吸引力法则公开课？',
            '邀请参加线下公开课的新朋友姓名', '邀请了多少位新朋友参加《与师父有约》落地读书会？', '邀请参加《与师父有约》落地读书会的新朋友姓名',
            '邀请了多少位新朋友参加吸引力法则90天线上践行班？', '邀请参加吸引力法则90天线上践行班的新朋友姓名', '投稿到简书专栏审核通过的文章名称',
            '颜色标记', '提交人', '修改人', '一、普通任务', '二、团队任务', '三、挑战任务', '来源', '填写设备', '操作系统',
            '浏览器', 'IP'
        ])
        # ## 异常值处理
        if data3['今日获得积分'].isna().any():
            replace_value = 0
            data3['今日获得积分'].fillna(replace_value, inplace=True)
        # print(data3['今日获得积分'].isna().any())

        # 战队内人员数
        teamInfo = pd.read_excel('data/teamInfo.xls')
        teamInfoList = np.array(teamInfo)
        group = data3['你所在的战队号']
        groupNum = group.drop_duplicates(keep='first', inplace=False)
        frame_list = []
        print('日打卡统计表：')
        for num in teamInfoList:
            for i in groupNum:
                data4 = data3[data3['你所在的战队号'] == i]
                if num[0] == i:
                    rate = Decimal(
                        len(data4) / num[1]).quantize(Decimal('0.00'))
                    integral = data4['今日获得积分']
                    sum_integral = sum(integral)
                    ave_integral = Decimal(
                        sum_integral / num[1]).quantize(Decimal('0.00'))
                    print('第', i, '战队:日应打卡人数:', num[1], ',日实际打卡人数:', len(data4), ',战队日打卡率:',
                          rate, ',战队日总积分:', sum_integral, ',战队日均积分:', ave_integral)
                    frame_list.append({
                        '1战队编号': i,
                        '2日应打卡人数': num[1],
                        '3日实际打卡人数': len(data4),
                        '4战队日打卡率': rate,
                        '5战队日总积分': sum_integral,
                        '6战队日均积分': ave_integral
                    })
        result = pd.DataFrame(frame_list)
        # ## 全班前三名
        sort = data3.sort_values("今日获得积分", ascending=False)
        sort = sort[(sort['你的战队编号'] != '助教') & (sort['你的战队编号'] != '教练')]
        sorces = sort['今日获得积分'].drop_duplicates(
            keep='first', inplace=False)[:3]
        rank_df = pd.DataFrame()
        rank = 1
        for sorce in sorces:
            sort_part = sort[sort['今日获得积分'] == sorce]
            sort_part['排名'] = rank
            # print('排名:', rank, sort_part[['你的战队编号', '你的姓名', '今日获得积分']])
            rank_df = rank_df.append(sort_part[['你的战队编号', '你的姓名', '今日获得积分', '排名']],
                                     ignore_index=True)
            rank += 1
        print('全班日积分前三名：')
        print(rank_df)
        # ## 战队前三名
        rank_sort = result.sort_values("5战队日总积分", ascending=False)
        rank_sorces = rank_sort['5战队日总积分'].drop_duplicates(
            keep='first', inplace=False)[:3]
        team_rank = 1
        team_rank_df = pd.DataFrame()
        for rank_sorce in rank_sorces:
            rank_sort_temp = rank_sort[rank_sort['5战队日总积分'] == rank_sorce]
            # print(rank_sort)
            rank_sort_temp['排名'] = team_rank
            team_rank_df = team_rank_df.append(rank_sort_temp)
            team_rank += 1
        print('战队日均积分前三名：')
        print(team_rank_df)
        # ## 将统计结果写入Excel
        writer = pd.ExcelWriter(output_path + yesterday + '打卡统计.xlsx')
        result.to_excel(excel_writer=writer,
                        encoding="utf_8_sig", sheet_name='打卡统计')
        rank_df.to_excel(excel_writer=writer,
                         encoding="utf_8_sig", sheet_name='全班日积分前三名')
        team_rank_df.to_excel(excel_writer=writer,
                              encoding="utf_8_sig",
                              sheet_name='战队日总积分前三名')
        writer.save()
        writer.close()

        create_png(result['1战队编号'],
                   result['6战队日均积分'], '战队编号', '积分', '战队日均积分排行榜', output_path, yesterday)
        create_png(result['1战队编号'],
                   result['4战队日打卡率'], '战队编号', '打卡率', '战队日打卡率排行榜', output_path, yesterday)
    else:
        print("文件不存在！")


# 生成柱状图
def create_png(x, y, xlabel, ylabel, title, output_path, time):
    plt.figure(figsize=(8, 6), dpi=80)
    # 柱子的宽度
    width = 0.5
    # 绘制柱状图, 每根柱子的颜色为紫罗兰色
    plt.bar(x,
            y,
            width,
            label="rainfall",
            color="#87CEFA")
    zhfont = matplotlib.font_manager.FontProperties(
        fname='simkai.ttf')
    # 设置横轴标签
    plt.xlabel(xlabel, fontproperties=zhfont, fontsize=16)
    # 设置纵轴标签
    plt.ylabel(ylabel, fontproperties=zhfont, fontsize=16)
    # 添加标题
    plt.title(title, fontproperties=zhfont, fontsize=20)
    # 添加纵横轴的刻度
    plt.xticks(range(0, len(x) + 1, 1), fontsize=10)  # 设置横坐标显示
    plt.yticks(fontsize=10)  # 设置纵坐标显示
    # 保存图片
    plt.savefig(output_path + time + title + '.png')
    # plt.show()


if __name__ == "__main__":
    print("程序开始*************************************************************************************************************")
    sratistics()
    print("程序结束*************************************************************************************************************")
