# 简单表格转化为latex

自己在写论文的时候发现想要把excel表格转化成latex好麻烦，
于是就想要写一个自动转化的软件

# 使用说明

xlsx文件可以转化为latex，但是xls文件由于我电脑上没有excel2003所以暂时不支持。

xlsx在仅有A列存在单元格合并的情况下可以正确转化为latex，其余情况不行
这主要是因为我不知道其余情况下的latex代码是什么样的

xls遇到的问题是，如果我用excel新版的office操作xls文档的话，xlrd库不能正确识别
文档中的合并单元格，所以暂时不支持。

# 开发计划

开发个屁

建议直接用 [TablesGenerator.com](https://www.tablesgenerator.com/)