Attribute VB_Name = "单工作簿转换维度"
Sub 单工作簿转换()

Dim r, c '数据源表格的最后非空单元格
Dim wb As Workbook, range2 As Worksheet, wb2 As Workbook
Dim month, row_num, year, m, t
'保养计划表wb '明细表range2
Application.ScreenUpdating = False
    Set wb2 = Workbooks.Add(xlWBATWorksheet)
    Set range2 = wb2.Sheets(1) '将汇总表赋值给range2
    'month = --Left(wb.Name, Application.Find("月", wb.Name, 1) - 1)
'调用当前表格
'获取当前表格的最大行数
    t = Timer
    row_num = 2 'range2.UsedRange.Rows.Count
    'range2.Rows(row_num & ":65536").Delete Shift:=xlShiftUp
    range2.Rows(row_num & ":65536").Clear '清空汇总表中源数据
    range2.Range("a1:e1") = Array("自编号", "保养级别", "保养日期", "保养场", "线路")
'    Dim filename
'    filename = Dir(wb2.Path & "\*.xls")
'    Do While filename <> "" '
'        If filename <> wb2.Name Then '判断文件是否是汇总数据工作簿
'            Set wb = GetObject(wb2.Path & "\" & filename) '将要汇总的工作簿赋值给变量wb
                '第一层for循环 遍历数据源的所有工作表sheets(x)
                Set wb = ThisWorkbook
                year = Left(wb.Sheets(1).Cells(2, 1), 4) + 1
                month = Mid(wb.Sheets(1).Cells(2, 1), 6, 2)
                For x = 1 To wb.Worksheets.Count
                    Dim location As String
                    '给字符串location赋值当前工作表的名字作为保养地点
                    location = wb.Sheets(x).Name
                    '第二层&第三层for循环，遍历当前工作表的所有单元格
                    For r = 5 To wb.Sheets(x).UsedRange.Rows.Count
                        For c = 3 To wb.Sheets(x).UsedRange.Columns.Count
                            Dim rng As Range
                            Set rng = wb.Sheets(x).Cells(r, c)
                            '判断当前单元格首字符是否等于“3”
                                        If Left(rng.Value, 1) = "3" Then
                                        'Exit For
                                        '不成立责判定当前单元格的行标题与列标题，即保养日期和保养级别
                                        
                            Dim xx, day, level As String, line As String
                            '判断活动单元格的列标题是否为空
                            If InStr(rng.Value, "路") > 0 Then
                                line = Mid(rng.Value, 7, InStr(rng.Value, "路") - 7)
                            Else
                                line = Right(rng.Value, 2)
                            End If
                            'line = Mid(rng.Value, 7, Application.Find("路", rng.Value) - 7) '获取线路号
                            num = rng.Column
                            xx = wb.Sheets(x).Cells(3, rng.Column).Value
                                While xx = ""
                                    num = num - 1
                                    xx = wb.Sheets(x).Cells(3, num).Value
                                Wend
                            level = Left(xx, 2) '保养级别赋值
                            '判断活动单元格的行标题
                            num = rng.Row
                            xx = wb.Sheets(x).Cells(rng.Row, 1).Value
                                While xx = ""
                                    num = num - 1
                                    xx = wb.Sheets(x).Cells(num, 1).Value
                                Wend
                            day = xx '保养日期赋值
                            
                            '对rang2输出单元格进行赋值
                            range2.Cells(row_num, 1) = Left(rng.Value, 5)
                            range2.Cells(row_num, 2) = level
                            range2.Cells(row_num, 3) = VBA.DateSerial(year, month, day)
                            range2.Cells(row_num, 4) = location
                            range2.Cells(row_num, 5) = line
                            row_num = row_num + 1
                            m = m + 1
                            End If
                        Next c
                    Next r
                Next x
'                wb.Close False
'            End If
'            filename = Dir '用dir函数取得其他文件名，并赋值给变量wb
'        Loop
    Columns.EntireColumn.AutoFit
    t = Timer - t
      wb2.sheet(1).Range("A1:E1").AutoFilter
    Application.ScreenUpdating = True
    MsgBox "完工" & Chr(10) & "搜集了" & m & "条保养信息呢" & Chr(10) & "只用了0" & t & "秒啦啦啦~", , "~\(≧▽≦)/~"
End Sub



