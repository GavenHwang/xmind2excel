import sys
import os.path

from src.excel2xmind import excel_to_xmind
from src.xmind2excel import xmind_to_excel

if __name__ == "__main__":
    # 获得请求参数
    input_file = sys.argv[1]
    input_path = os.path.join(os.path.dirname(__file__), input_file.strip())
    if input_path.endswith(".xlsx"):
        output_path = input_path.replace(".xlsx", ".xmind")
        excel_to_xmind(input_path, output_path)
    elif input_path.endswith(".xmind"):
        output_path = input_path.replace(".xmind", ".xlsx")
        xmind_to_excel(input_path, output_path)
    else:
        raise ValueError("请输入正确的文件格式：xlsx、xmind")
