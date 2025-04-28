# 1.简介
该脚本可实现excel格式与xmind格式测试用例之间的相互转换，默认仅支持禅道  
####  依赖说明：
需使用xmind8来编辑修改xmind文件，其他版本可能会导致乱码  
xmind8下载地址：https://xmind.app/download/xmind8/

# 2.安装
```shell
cd xmind2excel
pip install .
```

# 3.使用
```
e2m <测试用例文件>
```
如果测试用例文件为 .xlsx 格式，则将xlsx文件转为xmind文件  
如果测试用例文件为 .xmind 格式，则将xmind文件转为xlsx文件
