# -*- coding:utf-8 -*-
import xmind
import xlsxwriter
from datetime import datetime

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


def parse_xmind(xmind_file):
    workbook = xmind.load(xmind_file)
    sheet = workbook.getPrimarySheet()
    root_topic = sheet.getRootTopic()
    root_title = root_topic.getTitle()
    xmind_data = []
    for module_topic in root_topic.getSubTopics():
        if module_topic:
            module_title = module_topic.getTitle()
            for test_case_topic in module_topic.getSubTopics():
                if test_case_topic:
                    test_case_title = test_case_topic.getTitle()
                    row = {
                        "用例编号": "",
                        "所属产品": root_title,
                        "所属模块": module_title,
                        "用例标题": test_case_title,
                        "用例状态": "草稿",
                        "由谁创建": "",
                        "创建日期": datetime.strftime(datetime.now(), "%Y-%m-%d"),
                        "最后修改者": "",
                        "修改日期": datetime.strftime(datetime.now(), "%Y-%m-%d"),
                    }
                    for d in test_case_topic.getSubTopics():
                        if d:
                            d_title = d.getTitle()
                            # 处理"步骤和预期"
                            if d_title == '步骤和预期':
                                steps = []
                                expects = []
                                for i, step_sub_topic in enumerate(d.getSubTopics()):
                                    if step_sub_topic:
                                        # 处理"步骤"
                                        step_title = str(step_sub_topic.getTitle()).strip()
                                        if not step_title.startswith("%d." % (i + 1)):
                                            step_title = "%s.%s" % (i + 1, step_title)
                                        steps.append(step_title)
                                        # 处理"预期"
                                        expect_sub_topics = step_sub_topic.getSubTopics()
                                        if expect_sub_topics:
                                            expect_title = str(expect_sub_topics[0].getTitle()).strip()
                                            if not expect_title.startswith("%d." % (i + 1)):
                                                expect_title = "%s.%s" % (i + 1, expect_title)
                                            expects.append(expect_title)
                                row["步骤"] = "\n".join(steps)
                                row["预期"] = "\n".join(expects)
                            else:
                                sub_titles_text = []
                                for x in d.getSubTopics():
                                    if x:
                                        sub_titles_text.append(x.getTitle())
                                row[d_title] = "\n".join(sub_titles_text)
                            markers = test_case_topic.getMarkers()
                            row['优先级'] = markers[0].getMarkerId().name.split("-")[-1] if markers else "3"
                    xmind_data.append(row)
    return xmind_data


def xmind_to_excel(xmind_path, excel_path):
    xmind_data = parse_xmind(xmind_path)
    workbook = xlsxwriter.Workbook(excel_path)
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
        'font_size': 12,  # 字体大小
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
    for i in range(5, 9):
        worksheet.set_column(i, i, 48)
    workbook.close()
