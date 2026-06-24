Attribute VB_Name = "TOOL2输出更新对象"
Public Sub 输出更新对象()
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim folderPath As String: folderPath = wbDB.Path & "\"
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "净值数据浏览表\s+(\d{4}-\d{2}-\d{2})至(\d{4}-\d{2}-\d{2})\.xlsx$"
    
    Dim fileName As String, latestFile As String, latestEnd As String
    fileName = Dir(folderPath & "净值数据浏览表 *.xlsx")
    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Dim m As Object: Set m = regex.Execute(fileName)(0)
            If m.SubMatches(1) > latestEnd Then
                latestEnd = m.SubMatches(1)
                latestFile = fileName
            End If
        End If
        fileName = Dir()
    Loop
    
    If Len(latestFile) = 0 Then Exit Sub
    
    Debug.Print String(60, "=")
    Debug.Print "Excel版本: " & Application.Version
    Debug.Print "文件: " & latestFile
    Debug.Print "文件名结束日期: " & latestEnd
    
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(folderPath & latestFile, ReadOnly:=True)
    Dim wsSrc As Worksheet: Set wsSrc = wbSrc.Sheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row
    Debug.Print "数据末行: " & lastRow
    Debug.Print ""
    
    ' 看最后10行
    Debug.Print "末尾10行A/B列详情:"
    Dim r As Long
    For r = lastRow - 9 To lastRow
        If r >= 5 Then
            Dim aCell As Range: Set aCell = wsSrc.Cells(r, 1)
            Dim bCell As Range: Set bCell = wsSrc.Cells(r, 2)
            Debug.Print "  行" & r & _
                " | A.Value=[" & aCell.Value & "]" & _
                " | A.Text=[" & aCell.Text & "]" & _
                " | VarType=" & VarType(aCell.Value) & _
                " | IsDate=" & IsDate(aCell.Value) & _
                " | B=[" & bCell.Value & "]"
        End If
    Next r
    Debug.Print ""
    
    ' 统计源文件里每个日期出现多少次
    Debug.Print "源文件中各日期记录数:"
    Dim dateDict As Object: Set dateDict = CreateObject("Scripting.Dictionary")
    For r = 5 To lastRow
        Dim v As Variant: v = wsSrc.Cells(r, 1).Value
        Dim dStr As String
        If IsDate(v) Then
            dStr = Format(CDate(v), "yyyy-mm-dd")
        ElseIf Not IsEmpty(v) Then
            dStr = "<非日期:" & TypeName(v) & ":" & CStr(v) & ">"
        Else
            dStr = "<空>"
        End If
        
        If dateDict.Exists(dStr) Then
            dateDict(dStr) = dateDict(dStr) + 1
        Else
            dateDict(dStr) = 1
        End If
    Next r
    
    Dim k As Variant
    For Each k In dateDict.keys
        Debug.Print "  " & k & " : " & dateDict(k) & " 条"
    Next k
    
    wbSrc.Close SaveChanges:=False
    
    ' 同时查Sheet2(数据库)末日数据
    Debug.Print ""
    Debug.Print "Sheet2(数据库)各日期记录数(最近10个日期):"
    Set dateDict = CreateObject("Scripting.Dictionary")
    Dim wsDB As Worksheet: Set wsDB = wbDB.Sheets("Sheet2")
    Dim dbLast As Long: dbLast = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).row
    For r = 2 To dbLast
        v = wsDB.Cells(r, 1).Value
        If IsDate(v) Then
            dStr = Format(CDate(v), "yyyy-mm-dd")
            If dateDict.Exists(dStr) Then
                dateDict(dStr) = dateDict(dStr) + 1
            Else
                dateDict(dStr) = 1
            End If
        End If
    Next r
    
    ' 取最近10个日期
    Dim sortedDates() As String
    ReDim sortedDates(0 To dateDict.Count - 1)
    Dim idx As Long: idx = 0
    For Each k In dateDict.keys
        sortedDates(idx) = CStr(k)
        idx = idx + 1
    Next k
    ' 简单冒泡降序
    Dim i As Long, j As Long, tmp As String
    For i = 0 To UBound(sortedDates) - 1
        For j = 0 To UBound(sortedDates) - 1 - i
            If sortedDates(j) < sortedDates(j + 1) Then
                tmp = sortedDates(j)
                sortedDates(j) = sortedDates(j + 1)
                sortedDates(j + 1) = tmp
            End If
        Next j
    Next i
    
    Dim showCount As Long
    showCount = 10
    If UBound(sortedDates) + 1 < showCount Then showCount = UBound(sortedDates) + 1
    For i = 0 To showCount - 1
        Debug.Print "  " & sortedDates(i) & " : " & dateDict(sortedDates(i)) & " 条"
    Next i
End Sub

