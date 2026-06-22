Attribute VB_Name = "TOOL3批量删除数据"
Option Explicit

'==============================================================
' 模块: 按产品编号删除数据
' 功能: 用户输入产品编号 -> 预览统计 -> 二次确认 -> 删除Sheet2中所有该编号的记录
'==============================================================

Public Sub 批量删除数据()
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsData As Worksheet: Set wsData = wbDB.Sheets("Sheet2")
    
    '--- 1. 用户输入产品编号 ---
    Dim userInput As String
    userInput = InputBox("请输入要删除的产品编号:" & vbCrLf & vbCrLf & _
                         "提示: 该编号在Sheet2中的所有记录将被永久删除", _
                         "按产品编号删除数据")
    
    ' InputBox点取消返回空字符串
    If Len(userInput) = 0 Then
        Exit Sub
    End If
    
    Dim targetCode As String: targetCode = Trim(userInput)
    If Len(targetCode) = 0 Then
        MsgBox "产品编号不能为空。", vbExclamation
        Exit Sub
    End If
    
    '--- 2. 扫描Sheet2,统计匹配记录 ---
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    
    If lastRow < 2 Then
        MsgBox "Sheet2中没有数据。", vbInformation
        Exit Sub
    End If
    
    ' 读A、B列做匹配
    Dim dataArr As Variant
    dataArr = wsData.Range("A2:B" & lastRow).Value
    
    Dim matchCount As Long: matchCount = 0
    Dim minDate As Date: minDate = #1/1/9999#
    Dim maxDate As Date: maxDate = #1/1/1900#
    
    Dim i As Long
    Dim rowCode As String
    Dim rowDate As Variant
    Dim matchRows() As Long
    ReDim matchRows(1 To UBound(dataArr, 1))  ' 最多全部匹配
    
    For i = 1 To UBound(dataArr, 1)
        rowCode = Trim(CStr(dataArr(i, 2)))
        If rowCode = targetCode Then
            matchCount = matchCount + 1
            matchRows(matchCount) = i + 1  ' Sheet2中的实际行号
            
            rowDate = dataArr(i, 1)
            If IsDate(rowDate) Then
                Dim d As Date: d = CDate(rowDate)
                If d < minDate Then minDate = d
                If d > maxDate Then maxDate = d
            End If
        End If
    Next i
    
    If matchCount = 0 Then
        MsgBox "在Sheet2中未找到产品编号 [" & targetCode & "] 的记录。" & vbCrLf & vbCrLf & _
               "请检查编号是否正确(注意大小写和前后空格)。", vbInformation
        Exit Sub
    End If
    
    '--- 3. 二次确认 ---
    Dim msg As String
    msg = "找到产品编号 [" & targetCode & "] 的记录:" & vbCrLf & vbCrLf & _
          "  记录条数: " & matchCount & vbCrLf
    
    If maxDate >= minDate Then
        msg = msg & "  起始日期: " & Format(minDate, "yyyy-mm-dd") & vbCrLf & _
                    "  结束日期: " & Format(maxDate, "yyyy-mm-dd") & vbCrLf
    End If
    
    msg = msg & vbCrLf & _
          "确定要永久删除这些记录吗?" & vbCrLf & _
          "(此操作不可撤销)"
    
    Dim resp As VbMsgBoxResult
    resp = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, "确认删除")
    
    If resp <> vbYes Then
        MsgBox "已取消,未删除任何数据。", vbInformation
        Exit Sub
    End If
    
    '--- 4. 执行删除 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim t0 As Double: t0 = Timer
    
    ' 用Union构造一个包含所有目标行的Range,一次性删除
    Dim delRange As Range
    
    For i = 1 To matchCount
        If delRange Is Nothing Then
            Set delRange = wsData.Rows(matchRows(i))
        Else
            Set delRange = Union(delRange, wsData.Rows(matchRows(i)))
        End If
    Next i
    
    If Not delRange Is Nothing Then
        delRange.Delete Shift:=xlShiftUp
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    '--- 5. 完成提示 ---
    MsgBox "删除完成!" & vbCrLf & vbCrLf & _
           "产品编号: " & targetCode & vbCrLf & _
           "已删除: " & matchCount & " 条记录" & vbCrLf & _
           "耗时: " & Format(Timer - t0, "0.00") & " 秒", _
           vbInformation, "完成"
End Sub

