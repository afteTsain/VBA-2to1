Attribute VB_Name = "二维表脱水机By线路"
Sub Macro1()

Dim t
Dim drivername
Dim lines, xianlu As String
Dim iftrue
Dim xianlulist() As String
    t = Timer
    lines = "16/28/55/123/131/261/411/428/1098/B025"
    'lines = InputBox("输入线路名次以“/”分割 比如16/28/123", "开始执行", "16/28/55/123/131/261/411/428/1098/B025")
    If lines = "" Then
        MsgBox "你没有输入就点了确定"
        Exit Sub
    End If
    xianlulist = Split(lines, "/")
    For x = 1 To 2 ' ThisWorkbook.Sheets.Count
        Set myrange = Sheets(x).Range("B3").CurrentRegion '
        Sheets(x).Select
        Range("A1").Select
        Set myrange = ActiveCell.CurrentRegion
        
        row_num = myrange.Rows.Count
        col_num = myrange.Columns.Count
              
        For i = 4 To myrange.Rows.Count Step 1
            For j = 2 To myrange.Columns.Count Step 1
            
                iftrue = False
                'drivername = myrange.Offset(i, j).Resize(1, 1).Value
                drivername = Cells(i, j).Value
                If InStr(1, drivername, "路") <> 0 Then
                    xianlu = Mid(drivername, 7, InStr(1, drivername, "路") - 7)
                End If
                For t = 0 To UBound(xianlulist) Step 1
                    If xianlulist(t) = xianlu Then
                        iftrue = True
                    End If
                Next
                If iftrue = False Then
                    'myrange.Offset(i, j).Resize(1, 1).Value = ""
                    Cells(i, j).Value = ""
                End If
            
            Next
        Next
    Next
    t = Timer - t
    MsgBox Timer - t
End Sub
