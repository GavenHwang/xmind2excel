# 目录结构
```shell
├── README.md
├── main.py                           >>> 主程序入口
├── requirements.txt                  依赖包
├── src
│         ├── META-INF
│         │         └── manifest.xml  xmind修复文件
│         ├── excel2xmind.py          excel转xmind脚本
│         ├── templ                   模板文件
│         │         ├── testcase.xlsx
│         │         └── testcase.xmind
│         └── xmind2excel.py          xmind转excel脚本
```

# 使用说明
```
python main.py <用例文件>
如果用例文件为.xlsx格式，则将xlsx文件转为xmind文件
如果用例文件为.xmind格式，则将xmind文件转为xlsx文件
```
