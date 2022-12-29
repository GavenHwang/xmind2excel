# 目录结构
|----------src  
           |--xmind2excel.py 主程序文件  
           |--tmpl  
               |--testcase.xmind 测试用例模板文件  

# 查看脚本使用参数
### python xmind2excel.py -h
```
optional arguments:
  -h, --help     show this help message and exit
  -i INPUTFILE   指定xmind测试用例文件
  -o OUTPUTFILE  指定输出测试用例文件名，默认：测试用例.xlsx
  -u USER        用例作者（导入禅道时的用例创建者）
```

# 使用：
```
python xmind2excel.py -i testcase.xmind -o test.xlsx -u 小明
```

# 注意事项：
xmind模板文件必须使用xmind8工具来编辑，否则可能会出现解析乱码