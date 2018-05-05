# VBA
单工作簿转换维度，通过对目标表格插入模块的方式直接对单月保养计划进行操作，实现二维表转一维表的功能
-------------------------------------------------------------------------------------------

18-05-05 21:18  尝试了下增加自动筛选线路并复制粘贴生成新的工作表，然而却发现代码写得又臭又长，简直想删除这个模块，但是折腾了很久才弄好，用录制宏搞到的代码是这样子的    
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
    然而，我在eh上面找到的代码只用了一行“[a1].CurrentRegion.Sort key1:=[c1], order1:=1, Header:=1”就实现了关键字升序排序，感觉很厉害。
对了在测试的过程中发现用了好多并不需要的变量哦，有待优化ing
