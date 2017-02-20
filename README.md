# IntegrateExcel

## 介绍

用于将文件夹内所有Excel文件的第一张数据表内的第一行数据整合到同一个Excel中

- 本脚本支持xls和xlsx文件
- 本脚本运行需要perl环境



## 依赖包

本程序依赖：

- File::Basename
- Spreadsheet::XLSX
- Excel::Writer::XLSX
- Encode

有关依赖包的安装可以参考 [此链接](http://blog.csdn.net/memray/article/details/17543791)



## 使用方式

1. 将全部Excel表格拷贝进 `PutExcelsHere` 文件夹中
2. 运行脚本