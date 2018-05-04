# VBA
单工作簿转换维度，通过对目标表格插入模块的方式直接对单月保养计划进行操作，实现二维表转一维表的功能
range2.UsedRange.AutoFilter Field:=2, Criteria1:=Array( "10", "11", "123", "16", "26"), Operator:=xlFilterValues
Dim arr() As String
arr = Split(UserForm1.TextBox1.Text, "/")
