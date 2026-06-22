Attribute VB_Name = "导入净值数据"
Option Explicit

'==============================================================
' 模块: 导入净值数据
' 功能: 扫描当前文件夹下所有[净值数据浏览表 yyyy-mm-dd至yyyy-mm-dd.xlsx]
'       将数据按主键(日期+产品编号)增量合并到Sheet2
'==============================================================

Public Sub STEP1导入净值数据()
    
    Dim t0 As Double: t0 = Timer
    
    '--- 1. 准备环境 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsDB As Worksheet: Set wsDB = wbDB.Sheets("Sheet2")
    Dim folderPath As String: folderPath = wbDB.Path & "\"
    
    '--- 2. 扫描文件夹,正则匹配文件名,收集文件信息 ---
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "净值数据浏览表\s+(\d{4}-\d{2}-\d{2})至(\d{4}-\d{2}-\d{2})\.xlsx$"
    regex.IgnoreCase = True
    
    ' 用Dictionary存储: key=开始日期(用于排序), value=文件全路径
    Dim fileDict As Object
    Set fileDict = CreateObject("Scripting.Dictionary")
    
    Dim fileName As String
    fileName = Dir(folderPath & "净值数据浏览表 *.xlsx")
    
    Dim matches As Object, m As Object
    Dim dStart As Date, dEnd As Date
    Dim sortKey As String
    
    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Set matches = regex.Execute(fileName)
            Set m = matches(0)
            dStart = CDate(m.SubMatches(0))
            dEnd = CDate(m.SubMatches(1))
            
            If dStart <= dEnd Then
                ' 排序键: 开始日期+文件名,确保唯一且按日期升序
                sortKey = Format(dStart, "yyyy-mm-dd") & "|" & fileName
                fileDict.Add sortKey, folderPath & fileName
            End If
        End If
        fileName = Dir()
    Loop
    
    If fileDict.Count = 0 Then
        MsgBox "未找到任何[净值数据浏览表 yyyy-mm-dd至yyyy-mm-dd.xlsx]文件。" & vbCrLf & _
               "请检查当前文件夹: " & folderPath, vbExclamation, "提示"
        GoTo CleanUp
    End If
    
    ' 按key升序排序文件列表
    Dim sortedKeys() As String
    sortedKeys = SortKeys(fileDict.keys)
    
    '--- 3. 读取数据库Sheet2现有数据到内存,建立主键索引 ---
    Dim dbDict As Object
    Set dbDict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, lastCol As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).row
    lastCol = 10  ' A-J共10列
    
    Dim dbData As Variant
    Dim hasExistingData As Boolean
    
    If lastRow >= 2 Then
        hasExistingData = True
        dbData = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, lastCol)).Value
        
        Dim i As Long, pk As String
        For i = 1 To UBound(dbData, 1)
            pk = BuildKey(dbData(i, 1), dbData(i, 2))
            If Len(pk) > 0 Then
                ' value存储行号(在dbData数组中的索引)
                dbDict(pk) = i
            End If
        Next i
    Else
        hasExistingData = False
    End If
    
    '--- 4. 依次打开每个净值数据浏览表,合并数据 ---
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim srcLastRow As Long
    Dim srcData As Variant
    Dim newRows As Object  ' 存储新增行(主键不存在的)
    Set newRows = CreateObject("Scripting.Dictionary")
    
    Dim updatedCount As Long, insertedCount As Long
    updatedCount = 0
    insertedCount = 0
    
    Dim k As Variant, filePath As String
    For Each k In sortedKeys
        filePath = fileDict(k)
        
        On Error Resume Next
        Set wbSrc = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=0)
        If wbSrc Is Nothing Then
            MsgBox "无法打开文件: " & filePath & vbCrLf & "可能被占用,已跳过。", vbExclamation
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0
        
        Set wsSrc = wbSrc.Sheets("Sheet1")
        srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row
        
        If srcLastRow >= 5 Then
            srcData = wsSrc.Range(wsSrc.Cells(5, 1), wsSrc.Cells(srcLastRow, 10)).Value
            
            Dim j As Long, srcKey As String
            For j = 1 To UBound(srcData, 1)
                srcKey = BuildKey(srcData(j, 1), srcData(j, 2))
                If Len(srcKey) > 0 Then
                    If dbDict.Exists(srcKey) Then
                        ' 覆盖更新内存中的数据
                        Dim rowIdx As Long
                        rowIdx = dbDict(srcKey)
                        Dim c As Long
                        For c = 1 To 10
                            dbData(rowIdx, c) = srcData(j, c)
                        Next c
                        updatedCount = updatedCount + 1
                    Else
                        ' 新增行,先缓存
                        Dim newRow(1 To 10) As Variant
                        For c = 1 To 10
                            newRow(c) = srcData(j, c)
                        Next c
                        newRows(srcKey) = newRow
                        ' 同时加入dbDict,防止同批次重复
                        dbDict(srcKey) = -1  ' -1表示在newRows里
                        insertedCount = insertedCount + 1
                    End If
                End If
            Next j
        End If
        
        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
NextFile:
    Next k
    
    '--- 5. 一次性写回Sheet2 ---
    ' 5.1 先写回更新过的现有数据
    If hasExistingData And UBound(dbData, 1) > 0 Then
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(1 + UBound(dbData, 1), 10)).Value = dbData
    End If
    
    ' 5.2 再追加新增数据
    If newRows.Count > 0 Then
        Dim writeArr() As Variant
        ReDim writeArr(1 To newRows.Count, 1 To 10)
        Dim idx As Long: idx = 0
        Dim key As Variant
        For Each key In newRows.keys
            idx = idx + 1
            Dim arr As Variant: arr = newRows(key)
            For c = 1 To 10
                writeArr(idx, c) = arr(c)
            Next c
        Next key
        
        Dim writeStartRow As Long
        writeStartRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).row + 1
        If writeStartRow < 2 Then writeStartRow = 2
        
        wsDB.Range(wsDB.Cells(writeStartRow, 1), _
                   wsDB.Cells(writeStartRow + newRows.Count - 1, 10)).Value = writeArr
    End If
    
    ' 5.3 按A列日期升序排序(可选)
    Dim finalLastRow As Long
    finalLastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).row
    If finalLastRow >= 3 Then
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(finalLastRow, 10)).Sort _
            Key1:=wsDB.Cells(2, 1), Order1:=xlAscending, _
            Key2:=wsDB.Cells(2, 2), Order2:=xlAscending, _
            header:=xlNo
    End If
    
    MsgBox "导入完成!" & vbCrLf & _
           "扫描文件数: " & fileDict.Count & vbCrLf & _
           "覆盖更新: " & updatedCount & " 条" & vbCrLf & _
           "新增: " & insertedCount & " 条" & vbCrLf & _
           "耗时: " & Format(Timer - t0, "0.00") & " 秒", _
           vbInformation, "导入结果"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


'==============================================================
' 辅助函数: 构造主键 = yyyy-mm-dd|产品编号
'==============================================================
Private Function BuildKey(ByVal dateVal As Variant, ByVal codeVal As Variant) As String
    On Error Resume Next
    If IsEmpty(dateVal) Or IsNull(dateVal) Then Exit Function
    If Len(Trim(CStr(codeVal))) = 0 Then Exit Function
    
    Dim d As Date
    If IsDate(dateVal) Then
        d = CDate(dateVal)
        BuildKey = Format(d, "yyyy-mm-dd") & "|" & Trim(CStr(codeVal))
    Else
        ' 日期解析失败,用原始字符串
        BuildKey = Trim(CStr(dateVal)) & "|" & Trim(CStr(codeVal))
    End If
End Function


'==============================================================
' 辅助函数: 对字符串数组进行升序排序(简单冒泡,文件量不大够用)
'==============================================================
Private Function SortKeys(ByVal keys As Variant) As String()
    Dim arr() As String
    Dim n As Long: n = UBound(keys) - LBound(keys) + 1
    ReDim arr(0 To n - 1)
    
    Dim i As Long
    For i = 0 To n - 1
        arr(i) = CStr(keys(i))
    Next i
    
    Dim j As Long
    Dim tmp As String
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(j) > arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i
    
    SortKeys = arr
End Function


