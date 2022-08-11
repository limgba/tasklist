@echo off

rem 这是配置文件, 一般只需修改"文件路径"和"关注的名字"即可

rem 文件路径
set workbook="C:\Users\youyan\Desktop\python\data.xlsx"

rem 关注的名字列表
set follow_names="易聪, "

rem 显示excel中哪几列
set display_cols="1, 5, 6, 7, 8, 9, 3, 4"

rem 根据excel中第几列进行筛选, 筛选条件是空白, 注: 筛选列的数字必须是display_cols里面的其中一个
set filter_col=9

rem 程序窗口大小(参考: 桌面鼠标右键-显示设置-显示器分辨率)
set geometry="1280x540+0+0"

rem 注释, 可能在不同显示器用, 留着这条, 换显示器的时候把rem 去掉即可
rem set geometry="1280x540+0+0"

rem 一页显示多少条数据
set one_page_size=10

rem 字体大小
set front_size=20

rem 一列最大宽度(单位是像素)
set wraplength=500

rem 程序标题
set title="tasklist"

rem 关注的名字在哪几列(一般不用修改, 除非改了excel表结构)
set follow_cols="5, 6, 7, 8"

rem excel中有效数据的首行(一般不用修改, 除非改了excel表结构)
set first_row=3

