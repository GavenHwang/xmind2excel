import os.path
import sys

from e2m_package.excel2xmind import excel_to_xmind
from e2m_package.xmind2excel import xmind_to_excel


def main():
    # 获得请求参数
    if len(sys.argv) <= 1:
        raise ValueError("请指定一个测试用例文件，示例: e2m testcase.xlsx")
    elif len(sys.argv) > 2:
        raise ValueError("一次只能指定一个用例文件")
    input_file = sys.argv[1]
    input_path = os.path.abspath(input_file)
    if not os.path.exists(input_path):
        raise ValueError(f"{input_path}文件不存在！")
    if input_path.endswith(".xlsx"):
        output_path = input_path.replace(".xlsx", ".xmind")
        excel_to_xmind(input_path, output_path)
    elif input_path.endswith(".xmind"):
        output_path = input_path.replace(".xmind", ".xlsx")
        xmind_to_excel(input_path, output_path)
    else:
        raise ValueError("请输入正确的文件格式：xlsx、xmind")