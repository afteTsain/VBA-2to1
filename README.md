# VBA-2to1 Change log
单工作簿转换维度，通过对目标表格插入模块的方式直接对单月保养计划进行操作，实现二维表转一维表的功能
-------------------------------------------------------------------------------------------


## 4月4日至4月5日
* 完成最初的算法构思
* 手写算法结构顺序，确定需要申明的变量列表
* 完成初始代码敲打


## 2018-4-6 10:50
单文件遍历二转一调试通过

## 2018-4-7 13:00
增加三项代码需求，线路，年份，和多文件遍历

## 2018-4-7 14:05
关键字route/line提取失败，Application.Find不能处理错误值

## 2018-4-7 14:08
修正worksheet.count---》wb.count

## 2018-4-7 14:14
通过提取cells（3，2）并赋值year，二周目提取月份关键字month

## 2018-4-7 14:35
最初是单元格查找函数替换成字符串替换函数instr，line关键字提取达成

## 2018-4-7 15:39
遍历当前文件夹下所有表格汇总达成(*.xlsx)

## 2018-4-7 16:30
增加计数器变量m，并以msgbox提示代码运行结束

## 2018-4-8 10:20
"Workbooks(""保养明细.xlsm"")——》ThisWorkbook
关键字容错性++，昨晚上看视频学的"

## 2018-4-9 09:30
统计2016年的保养数据

## 2018-4-28 11:00
单工作簿转换维度，通过对目标表格插入模块的方式直接对单月保养计划进行操作

## 2018-4-28 11:03
"Sheet1.UsedRange.Columns.Autofit，对表格进行匹配列宽
通过arry对数组对进行标题栏赋值"

## 2018-4-29 00:00
加入搜集msgbox返回值逻辑语句，如果返回vbNo 则1.清空单元格，2.关闭新建文件，其实可以直接注释掉第一步直接执行第二步就好了

## 2018-4-30 00:00
用数组赋值输出区域，提升代码运行效率，大约节约了40%左右的时间range2.Cells(row_num, 1).Resize(1, 5) = Array(Left(rng.Value, 5), level, VBA.DateSerial(year, month, day), location, line)

## 2018-5-4 13:04
form""本数据仅供参考，不做实际生产计划执行""to "本表数据仅供参考，不保证数据100%准确,以上"

## 2018-5-4 16:44
保养级别提取关键字to4，调整数据输出顺序，添加“星期几”关键字，msgbox语句整体较长的一行代码换行书写




## 18-05-05 21:18  
尝试了下增加自动筛选线路并复制粘贴生成新的工作表，然而却发现代码写得又臭又长，简直想删除这个模块，但是折腾了很久才弄好，用录制宏搞到的代码是这样子的    
ActiveWorkbook.Worksheets("车队").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("车队").Sort.SortFields.Add Key:=Range("C1"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("车队").Sort
        .SetRange Range("A2:F160")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
> 然而，我在eh上面找到的代码只用了一行“[a1].CurrentRegion.Sort key1:=[c1], order1:=1, Header:=1”就实现了关键字升序排序，感觉很厉害。
对了在测试的过程中发现用了好多并不需要的变量哦，有待优化ing
