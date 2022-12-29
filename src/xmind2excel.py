# -*- coding:utf-8 -*-
from datetime import datetime
import xlsxwriter
import xmind
import argparse
import os

sub_titles = [
    '用例编号',
    '所属产品',
    '分支',
    '所属模块',
    '相关需求',
    '用例标题',
    '前置条件',
    '步骤',
    '预期',
    '实际情况',
    '关键词',
    '优先级',
    '用例类型',
    '适用阶段',
    '用例状态',
    'B',
    'R',
    'S',
    '结果',
    '由谁创建',
    '创建日期',
    '最后修改者',
    '修改日期',
    '用例版本',
    '相关用例',
    '子状态',
    '附件'
]


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', type=str, dest='inputfile', default="测试用例.xmind", help='指定xmind测试用例文件')
    parser.add_argument('-o', type=str, dest='outputfile', default="测试用例.xlsx", help='指定输出测试用例文件名，默认：测试用例.xlsx')
    parser.add_argument('-u', type=str, dest='user', help='用例作者')
    args = parser.parse_args()
    # 参数为空，直接退出
    if args.inputfile is None:
        parser.print_help()
        exit()
    return args


def parse_xmind(xmind_file, user):
    workbook = xmind.load(xmind_file)
    sheet = workbook.getPrimarySheet()
    root_topic = sheet.getRootTopic()
    root = root_topic.getData()
    xmind_data = []
    for module in root['topics']:
        for test_case in module['topics']:
            row = {
                "用例编号": "",
                "所属产品": root['title'],
                "所属模块": module['title'],
                "用例标题": test_case['title'],
                "用例状态": "草稿",
                "由谁创建": user,
                "创建日期": datetime.strftime(datetime.now(), "%Y-%m-%d"),
                "最后修改者": user,
                "修改日期": datetime.strftime(datetime.now(), "%Y-%m-%d"),
            }
            for d in test_case['topics']:
                # 处理"步骤和预期"
                if d["title"] == '步骤和预期':
                    steps = []
                    expects = []
                    for i in range(len(d['topics'])):
                        # 处理"步骤"
                        step_title = str(d['topics'][i]["title"]).strip()
                        if not step_title.startswith("%d." % (i + 1)):
                            step_title = "%s.%s" % (i + 1, step_title)
                        steps.append(step_title)
                        # 处理"预期"
                        if d['topics'][i].get("topics"):
                            expect_title = str(d['topics'][i]["topics"][0]["title"]).strip()
                            if not expect_title.startswith("%d." % (i + 1)):
                                expect_title = "%s.%s" % (i + 1, expect_title)
                            expects.append(expect_title)
                    row["步骤"] = "\n".join(steps)
                    row["预期"] = "\n".join(expects)
                else:
                    row[d['title']] = "\n".join([x['title'] for x in d['topics']])
                row['优先级'] = test_case['markers'][0] and test_case['markers'][0].replace("priority-", "") or "3"
            xmind_data.append(row)
    return xmind_data


def write_excel(xmind_data, outputfile):
    workbook = xlsxwriter.Workbook(outputfile)
    # 表头样式
    sub_title_style = workbook.add_format({
        'font_name': "微软雅黑",  # 字体
        'font_size': 14,  # 字体大小
        'font_color': "#ffffff",  # 字体大小
        'bold': True,  # 字体加粗
        'border': 1,  # 单元格边框宽度
        'align': 'center',  # 左右对齐方式
        'valign': 'vcenter',  # 上下对齐方式
        'fg_color': '#292f8b',  # 背景色
        'num_format': 0,  # 数字、日期格式化  '￥#,##0.00'、'yyyy-m-d h:mm:ss'
    })
    sub_title_style.set_text_wrap()
    # 表体样式
    text_style = workbook.add_format({
        'font_name': "微软雅黑",  # 字体
        'font_size': 14,  # 字体大小
        'bold': False,  # 字体加粗
        'border': 1,  # 单元格边框宽度
        'align': 'left',  # 左右对齐方式
        'valign': 'vcenter',  # 上下对齐方式
        'fg_color': '#a8d2e6',  # 背景色
        'num_format': 0,  # 数字、日期格式化  '￥#,##0.00'、'yyyy-m-d h:mm:ss'
    })
    text_style.set_text_wrap()
    # 新加一个sheet页
    worksheet = workbook.add_worksheet('用例')
    # 添加表头
    for i in range(len(sub_titles)):
        worksheet.write(0, i, sub_titles[i], sub_title_style)
    row = 1
    for data in xmind_data:
        for i in range(len(sub_titles)):
            worksheet.write(row, i, data.get(sub_titles[i], ""), text_style)
        row += 1
    for i in range(len(sub_titles)):
        worksheet.set_column(i, i, 12)
    workbook.close()


if __name__ == "__main__":
    # 获得请求参数
    args = get_args()
    inputfile = os.getcwd() + os.path.sep + args.inputfile
    outputfile = os.getcwd() + os.path.sep + args.outputfile
    if not os.path.isfile(inputfile):
        raise ValueError("输入文件不存在！")
    xmind_data = parse_xmind(inputfile, args.user)
    write_excel(xmind_data, outputfile)
    print('成功转换%d条用例\n用例生成位置：%s' % (len(xmind_data), outputfile))
