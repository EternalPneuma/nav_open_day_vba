Option Explicit
Private g_fillCounts As Object

'==============================================================
' 模块: 导出产品数据
' 功能: 按维度表中的产品,逐个生成sheet到一个新xlsx文件
'       每个sheet以产品简称命名,包含该产品全部历史数据,按净值日期升序
'==============================================================

Public Sub STEP3导出产品数据()
    
    Dim t0 As Double: t0 = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsDim As Worksheet: Set wsDim = wbDB.Sheets("Sheet1")
    Dim wsData As Worksheet: Set wsData = wbDB.Sheets("Sheet2")
    
    ' 初始化填补统计收集器
    Set g_fillCounts = CreateObject("Scripting.Dictionary")
    
    '--- 1. 读取维度表: 产品编号 -> 产品简称 ---
    Dim dimDict As Object
    Set dimDict = CreateObject("Scripting.Dictionary")
    
    Dim dimLastRow As Long
    dimLastRow = wsDim.Cells(wsDim.Rows.Count, "A").End(xlUp).row
    
    If dimLastRow < 2 Then
        MsgBox "维度表(Sheet1)中没有数据,请先完善维度表。", vbExclamation
        GoTo CleanUp
    End If
    
    Dim dimArr As Variant
    dimArr = wsDim.Range("A2:C" & dimLastRow).Value  ' A=产品编号,B=产品名称,C=产品简称
    
    Dim i As Long, prodCode As String, prodShort As String
    For i = 1 To UBound(dimArr, 1)
        prodCode = Trim(CStr(dimArr(i, 1)))
        prodShort = Trim(CStr(dimArr(i, 3)))
        If Len(prodCode) > 0 And Len(prodShort) > 0 Then
            dimDict(prodCode) = prodShort
        End If
    Next i
    
    If dimDict.Count = 0 Then
        MsgBox "维度表中未读取到有效的产品编号和简称。", vbExclamation
        GoTo CleanUp
    End If
    
    '--- 2. 读取Sheet2全部数据,按产品编号分组,同时去重 ---
    Dim dataLastRow As Long
    dataLastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    
    If dataLastRow < 2 Then
        MsgBox "Sheet2中没有数据可导出。", vbExclamation
        GoTo CleanUp
    End If
    
    ' 读取表头(第1行A-J)
    Dim header As Variant
    header = wsData.Range("A1:J1").Value
    
    ' 读取数据
    Dim dataArr As Variant
    dataArr = wsData.Range("A2:J" & dataLastRow).Value
    
    ' 按产品编号分组: groupDict(prodCode) = 子Dictionary( pk -> 行数组 )
    ' 同时用pk(日期+编号)做去重,后出现的覆盖先出现的
    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")
    
    Dim maxDate As Date: maxDate = 0
    
    Dim j As Long, c As Long
    Dim curCode As String, pk As String
    Dim curDate As Variant
    
    For j = 1 To UBound(dataArr, 1)
        curDate = dataArr(j, 1)
        curCode = Trim(CStr(dataArr(j, 2)))
        
        If Len(curCode) = 0 Then GoTo NextRow
        If IsEmpty(curDate) Or Not IsDate(curDate) Then GoTo NextRow
        
        ' 跟踪最大日期(用于文件名)
        If CDate(curDate) > maxDate Then maxDate = CDate(curDate)
        
        ' 主键
        pk = Format(CDate(curDate), "yyyy-mm-dd") & "|" & curCode
        
        ' 取出/创建该产品的子Dictionary
        Dim subDict As Object
        If groupDict.Exists(curCode) Then
            Set subDict = groupDict(curCode)
        Else
            Set subDict = CreateObject("Scripting.Dictionary")
            groupDict.Add curCode, subDict
        End If
        
        ' 缓存该行(数组形式)
        Dim rowArr(1 To 10) As Variant
        For c = 1 To 10
            rowArr(c) = dataArr(j, c)
        Next c
        subDict(pk) = rowArr  ' 重复主键自动覆盖
        
NextRow:
    Next j
    
    '--- 3. 创建新工作簿,逐个产品生成sheet ---
    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add
    
    ' 删除默认创建的多余sheet,只留一个占位
    Do While wbOut.Sheets.Count > 1
        wbOut.Sheets(wbOut.Sheets.Count).Delete
    Loop
    
    Dim usedSheetNames As Object
    Set usedSheetNames = CreateObject("Scripting.Dictionary")
    
    Dim exportedCount As Long: exportedCount = 0
    Dim missingProducts As String: missingProducts = ""
    Dim emptyProducts As String: emptyProducts = ""
    
    Dim isFirstSheet As Boolean: isFirstSheet = True
    Dim wsOut As Worksheet
    
    ' 按维度表的顺序遍历,保证sheet顺序与维度表一致
    Dim codeKey As Variant
    For Each codeKey In dimDict.keys
        prodCode = CStr(codeKey)
        prodShort = dimDict(prodCode)
        
        ' 检查该产品是否有数据
        If Not groupDict.Exists(prodCode) Then
            emptyProducts = emptyProducts & prodCode & "(" & prodShort & "), "
            GoTo NextProduct
        End If
        
        Set subDict = groupDict(prodCode)
        If subDict.Count = 0 Then
            emptyProducts = emptyProducts & prodCode & "(" & prodShort & "), "
            GoTo NextProduct
        End If
        
        ' 清洗sheet名
        Dim cleanName As String
        cleanName = CleanSheetName(prodShort, usedSheetNames)
        usedSheetNames(cleanName) = 1
        
        ' 创建sheet
        If isFirstSheet Then
            Set wsOut = wbOut.Sheets(1)
            wsOut.Name = cleanName
            isFirstSheet = False
        Else
            Set wsOut = wbOut.Sheets.Add(After:=wbOut.Sheets(wbOut.Sheets.Count))
            wsOut.Name = cleanName
        End If
        
        ' 写入表头
        wsOut.Range("A1:J1").Value = header
        
        ' 把subDict里的数据按日期升序排列后写入
        Dim writeArr() As Variant
        Dim nRows As Long: nRows = subDict.Count
        ReDim writeArr(1 To nRows, 1 To 10)
        
        ' 先把所有行收集到临时数组(包含排序键)
        Dim tmpArr() As Variant
        ReDim tmpArr(1 To nRows, 1 To 11)  ' 第11列存排序键
        
        Dim idx As Long: idx = 0
        Dim pkKey As Variant
        For Each pkKey In subDict.keys
            idx = idx + 1
            Dim arr As Variant: arr = subDict(pkKey)
            For c = 1 To 10
                tmpArr(idx, c) = arr(c)
            Next c
            tmpArr(idx, 11) = CStr(pkKey)  ' 主键作为排序键(yyyy-mm-dd|code)
        Next pkKey
        
        ' 按第11列升序排序(冒泡,产品内数据量一般不大)
        SortByCol tmpArr, 11
        
       ' 对净值列(第4列)做分红日平滑,覆盖原值
        SmoothDividendDays tmpArr, 4
        
        ' 基于平滑后的净值,重算30日年化收益率,覆盖第7列(G列)
        Calc30DayAnnualYield tmpArr, 1, 4, 7, prodCode
        
        ' 拷贝到writeArr
        For idx = 1 To nRows
            For c = 1 To 10
                writeArr(idx, c) = tmpArr(idx, c)
            Next c
        Next idx
        
        ' 一次性写入数据区
        wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(nRows + 1, 10)).Value = writeArr
        
        ' 简单格式化: 表头加粗,日期列格式
        wsOut.Range("A1:J1").Font.Bold = True
        wsOut.Range("A:A").NumberFormat = "yyyy-mm-dd"
        wsOut.Columns("A:J").AutoFit
        
        exportedCount = exportedCount + 1
        
NextProduct:
    Next codeKey
    
    ' 检查Sheet2中存在但维度表中没有的产品
    Dim dataCodeKey As Variant
    For Each dataCodeKey In groupDict.keys
        If Not dimDict.Exists(CStr(dataCodeKey)) Then
            missingProducts = missingProducts & CStr(dataCodeKey) & ", "
        End If
    Next dataCodeKey
    
    '--- 4. 保存文件 ---
    If exportedCount = 0 Then
        wbOut.Close SaveChanges:=False
        MsgBox "没有可导出的产品数据。", vbExclamation
        GoTo CleanUp
    End If
    
    Dim outFileName As String
    If maxDate > 0 Then
        outFileName = "产品净值汇总_" & Format(maxDate, "yyyymmdd") & ".xlsx"
    Else
        outFileName = "产品净值汇总_" & Format(Now, "yyyymmdd") & ".xlsx"
    End If
    
    Dim outPath As String
    outPath = wbDB.Path & "\" & outFileName
    
    ' 如果同名文件已存在,先删除(避免SaveAs弹窗)
    On Error Resume Next
    If Dir(outPath) <> "" Then Kill outPath
    On Error GoTo 0
    
    wbOut.SaveAs fileName:=outPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    
    '--- 5. 汇总提示 ---
    Dim msg As String
    msg = "导出完成!" & vbCrLf & _
          "输出文件: " & outFileName & vbCrLf & _
          "导出产品数: " & exportedCount & " 个" & vbCrLf & _
          "耗时: " & Format(Timer - t0, "0.00") & " 秒"
    
    If Len(emptyProducts) > 0 Then
        If Len(emptyProducts) > 2 Then emptyProducts = Left(emptyProducts, Len(emptyProducts) - 2)
        msg = msg & vbCrLf & vbCrLf & "维度表中存在但Sheet2无数据(已跳过):" & vbCrLf & emptyProducts
    End If
    
    If Len(missingProducts) > 0 Then
        If Len(missingProducts) > 2 Then missingProducts = Left(missingProducts, Len(missingProducts) - 2)
        msg = msg & vbCrLf & vbCrLf & "Sheet2中存在但维度表未配置(已跳过):" & vbCrLf & missingProducts
    End If
    
    Dim totalFilled As Long: totalFilled = 0
    Dim kFill As Variant
    For Each kFill In g_fillCounts.keys
        totalFilled = totalFilled + g_fillCounts(kFill)
    Next kFill
    
    WriteFillCountsToSheet3 wbDB, g_fillCounts
    
    If totalFilled > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "30日年化收益率: 计算填补 " & totalFilled & " 条" & vbCrLf & _
              "(详情见Sheet3 M列;原始有效值已保留)"
    End If
    
    MsgBox msg, vbInformation, "导出结果"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


'==============================================================
' 辅助函数: 清洗sheet名
'  - 移除非法字符 \ / ? * [ ] :
'  - 长度截断到31
'  - 重名自动加后缀
'==============================================================
Private Function CleanSheetName(ByVal rawName As String, ByVal usedDict As Object) As String
    Dim s As String: s = rawName
    
    ' 替换非法字符
    Dim badChars As Variant
    badChars = Array("\", "/", "?", "*", "[", "]", ":")
    Dim k As Long
    For k = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(k), "-")
    Next k
    
    ' 去除首尾单引号(Excel不允许sheet名以单引号开头或结尾)
    Do While Left(s, 1) = "'"
        s = Mid(s, 2)
    Loop
    Do While Right(s, 1) = "'"
        s = Left(s, Len(s) - 1)
    Loop
    
    s = Trim(s)
    If Len(s) = 0 Then s = "未命名"
    
    ' 截断到31字符
    If Len(s) > 31 Then s = Left(s, 31)
    
    ' 处理重名
    Dim baseName As String: baseName = s
    Dim suffix As Long: suffix = 2
    Do While usedDict.Exists(s)
        Dim suffixStr As String: suffixStr = "_" & suffix
        If Len(baseName) + Len(suffixStr) > 31 Then
            s = Left(baseName, 31 - Len(suffixStr)) & suffixStr
        Else
            s = baseName & suffixStr
        End If
        suffix = suffix + 1
    Loop
    
    CleanSheetName = s
End Function


'==============================================================
' 辅助函数: 按指定列对二维数组升序排序(冒泡)
'==============================================================
Private Sub SortByCol(ByRef arr As Variant, ByVal sortCol As Long)
    Dim n As Long: n = UBound(arr, 1)
    Dim nCols As Long: nCols = UBound(arr, 2)
    
    Dim i As Long, j As Long, c As Long
    Dim tmp As Variant
    
    For i = 1 To n - 1
        For j = 1 To n - i
            If CStr(arr(j, sortCol)) > CStr(arr(j + 1, sortCol)) Then
                For c = 1 To nCols
                    tmp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

'==============================================================
' 辅助过程: 对二维数组的指定列做"分红日平滑"
'   - 算法: 基于"反弹特征"识别孤立异常点
'   - 条件: |V(D) - (V(D-1)+V(D+1))/2| > MAX( |V(D+1)-V(D-1)|*K, V(D-1)*threshold )
'   - 替换: 异常点V(D) = (V(D-1)+V(D+1))/2
'   - 边界: 首尾点不处理
'   - 调用前提: 数组已按日期升序排序
'==============================================================
Private Sub SmoothDividendDays(ByRef arr As Variant, ByVal navCol As Long)
    Const K_RATIO As Double = 3#      ' 偏离倍数
    Const THRESHOLD As Double = 0.0005 ' 绝对阈值: 0.05%
    
    Dim n As Long: n = UBound(arr, 1)
    If n < 3 Then Exit Sub  ' 少于3个点无法平滑
    
    ' 收集有效数值的索引(跳过非数值/空值)
    ' 注意: 平滑只在"连续有效"的位置之间进行
    ' 即如果 arr(i, navCol) 是数值, 才参与判定
    
    Dim i As Long
    Dim prevV As Double, curV As Double, nextV As Double
    Dim expected As Double, jump As Double, baseline As Double
    Dim limit As Double
    
    ' 用一个新数组保存平滑后的净值,避免边平滑边判定(否则后一个点会基于已修改的前一个点判定)
    Dim newNav() As Double
    ReDim newNav(1 To n)
    Dim isValid() As Boolean
    ReDim isValid(1 To n)
    
    For i = 1 To n
        If IsNumeric(arr(i, navCol)) And Not IsEmpty(arr(i, navCol)) Then
            newNav(i) = CDbl(arr(i, navCol))
            isValid(i) = True
        Else
            isValid(i) = False
        End If
    Next i
    
    ' 对内部点(2 到 n-1)做判定,使用ORIGINAL值判断而非已修改值
    For i = 2 To n - 1
        If Not isValid(i) Then GoTo NextPoint
        If Not isValid(i - 1) Then GoTo NextPoint
        If Not isValid(i + 1) Then GoTo NextPoint
        
        prevV = CDbl(arr(i - 1, navCol))
        curV = CDbl(arr(i, navCol))
        nextV = CDbl(arr(i + 1, navCol))
        
        expected = (prevV + nextV) / 2
        jump = Abs(curV - expected)
        baseline = Abs(nextV - prevV)
        
        ' 判定阈值: max(baseline*K, prev*threshold)
        limit = baseline * K_RATIO
        If prevV * THRESHOLD > limit Then limit = prevV * THRESHOLD
        
        If jump > limit Then
            ' 异常点,替换
            newNav(i) = expected
        End If
NextPoint:
    Next i
    
    ' 把newNav写回arr
    For i = 1 To n
        If isValid(i) Then
            arr(i, navCol) = newNav(i)
        End If
    Next i
End Sub

'==============================================================
' 辅助过程: 基于净值序列重算30日年化收益率(复利公式,365天年)
'   - 公式: ((V_t / V_t-30) ^ (365 / 实际间隔) - 1) × 100
'   - 匹配规则: 找日期-30,失败则向前匹配-31/-32/-33,最多4次
'   - 合并逻辑:
'       * 原始值有效(非空/非""/非0/非#N/A)且与计算值偏离<=阈值 → 保留原始值
'       * 原始值有效但偏离>阈值                              → 用计算值覆盖
'       * 原始值无效                                          → 用计算值填补
'       * 计算失败时,如果原始值有效则保留,否则写#N/A
'
' 参数:
'   arr        - 二维数组(已按日期升序排序)
'   dateCol    - 日期列
'   navCol     - 净值列
'   yieldCol   - 年化收益列
'   prodCode   - 当前产品编号(用于记录填补统计)
'==============================================================
Private Sub Calc30DayAnnualYield(ByRef arr As Variant, _
                                  ByVal dateCol As Long, _
                                  ByVal navCol As Long, _
                                  ByVal yieldCol As Long, _
                                  ByVal prodCode As String)
    Const TARGET_DAYS As Long = 30
    Const MAX_LOOKBACK As Long = 3
    Const DAYS_PER_YEAR As Double = 365#
    Const DEVIATION_THRESHOLD As Double = 0.5   ' 偏离阈值(百分点): 原始与计算差>1.0则用计算值
    Const FLAT_NAV_THRESHOLD As Double = 0.0001  ' 净值波动<1基点视为"不变"
    
    Dim n As Long: n = UBound(arr, 1)
    If n < 2 Then Exit Sub
    
    '--- 0. 检测是否为"净值不变型"产品(净值=固定值,收益靠分红) ---
    Dim navMin As Double, navMax As Double
    Dim hasFirstNav As Boolean: hasFirstNav = False
    Dim ii As Long
    For ii = 1 To n
        If IsNumeric(arr(ii, navCol)) And Not IsEmpty(arr(ii, navCol)) Then
            Dim navVal As Double: navVal = CDbl(arr(ii, navCol))
            If navVal > 0 Then
                If Not hasFirstNav Then
                    navMin = navVal
                    navMax = navVal
                    hasFirstNav = True
                Else
                    If navVal < navMin Then navMin = navVal
                    If navVal > navMax Then navMax = navVal
                End If
            End If
        End If
    Next ii
    
    Dim isFlatNav As Boolean: isFlatNav = False
    If hasFirstNav Then
        If (navMax - navMin) < FLAT_NAV_THRESHOLD Then
            isFlatNav = True
        End If
    End If
    
    ' 净值不变型: 完全跳过计算,保留原始G列值不变
    ' (该产品的填补数记0)
    If isFlatNav Then
        If Not g_fillCounts Is Nothing And Len(prodCode) > 0 Then
            g_fillCounts(prodCode) = 0
        End If
        Exit Sub
    End If
    
    '--- 1. 建立 日期(Long) -> 净值 的字典索引 ---
    Dim navDict As Object
    Set navDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim d As Date
    
    For i = 1 To n
        If IsDate(arr(i, dateCol)) And IsNumeric(arr(i, navCol)) Then
            d = CDate(arr(i, dateCol))
            If CDbl(arr(i, navCol)) > 0 Then
                navDict(CLng(CDbl(d))) = CDbl(arr(i, navCol))
            End If
        End If
    Next i
    
    '--- 2. 逐行计算并合并 ---
    Dim curDate As Date, curNav As Double
    Dim baseDateNum As Long, baseNav As Double
    Dim offset As Long, actualGap As Long
    Dim foundBase As Boolean
    Dim ratio As Double, calcYield As Double
    Dim hasCalc As Boolean
    Dim rawVal As Variant, rawValid As Boolean, rawNum As Double
    Dim filledCount As Long: filledCount = 0
    
    For i = 1 To n
        '--- 2.1 判断原始值是否"有效" ---
        ' 有效定义: 非空 + 非"" + 非0 + 非#N/A错误
        rawValid = False
        rawNum = 0
        rawVal = arr(i, yieldCol)
        
        If Not IsError(rawVal) Then
            If Not IsEmpty(rawVal) Then
                If IsNumeric(rawVal) Then
                    rawNum = CDbl(rawVal)
                    If rawNum <> 0 Then
                        rawValid = True
                    End If
                End If
            End If
        End If
        ' IsError、IsEmpty、非数值、""、0 都视为无效
        
        '--- 2.2 尝试计算理论值 ---
        hasCalc = False
        calcYield = 0
        
        ' 当前行净值有效才能算
        If IsDate(arr(i, dateCol)) And IsNumeric(arr(i, navCol)) Then
            If CDbl(arr(i, navCol)) > 0 Then
                curDate = CDate(arr(i, dateCol))
                curNav = CDbl(arr(i, navCol))
                
                ' 查找 D-30, D-31, D-32, D-33
                foundBase = False
                For offset = 0 To MAX_LOOKBACK
                    baseDateNum = CLng(CDbl(curDate)) - TARGET_DAYS - offset
                    If navDict.Exists(baseDateNum) Then
                        baseNav = navDict(baseDateNum)
                        actualGap = TARGET_DAYS + offset
                        foundBase = True
                        Exit For
                    End If
                Next offset
                
                If foundBase Then
                    ratio = curNav / baseNav
                    If ratio > 0 Then
                        calcYield = (ratio ^ (DAYS_PER_YEAR / actualGap) - 1) * 100
                        hasCalc = True
                    End If
                End If
            End If
        End If
        
        '--- 2.3 合并逻辑 ---
        If rawValid And hasCalc Then
            ' 都有: 看偏离
            If Abs(rawNum - calcYield) > DEVIATION_THRESHOLD Then
                ' 偏离过大,用计算值覆盖
                arr(i, yieldCol) = calcYield
                filledCount = filledCount + 1
            Else
                ' 偏离合理,保留原始
                arr(i, yieldCol) = rawNum
            End If
        ElseIf rawValid And Not hasCalc Then
            ' 只有原始: 保留
            arr(i, yieldCol) = rawNum
        ElseIf Not rawValid And hasCalc Then
            ' 只有计算: 填补
            arr(i, yieldCol) = calcYield
            filledCount = filledCount + 1
        Else
            ' 都没有: #N/A
            arr(i, yieldCol) = CVErr(xlErrNA)
        End If
    Next i
    
    '--- 3. 记录该产品的填补条数 ---
    If Not g_fillCounts Is Nothing And Len(prodCode) > 0 Then
        g_fillCounts(prodCode) = filledCount
    End If
End Sub
'==============================================================
' 辅助过程: 把每个产品的填补条数写到Sheet3的M列
'   - 如果Sheet3不存在,直接返回(不报错)
'   - 按A列产品编号匹配Sheet3的行
'   - M列表头写"计算填补条数"
'==============================================================
Private Sub WriteFillCountsToSheet3(ByVal wbDB As Workbook, ByVal fillDict As Object)
    Dim wsRpt As Worksheet
    
    On Error Resume Next
    Set wsRpt = wbDB.Sheets("Sheet3")
    On Error GoTo 0
    
    If wsRpt Is Nothing Then Exit Sub
    If fillDict Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).row
    
    If lastRow < 2 Then Exit Sub  ' Sheet3没有数据行
    
    ' 写表头
    wsRpt.Cells(1, 13).Value = "计算填补条数"
    wsRpt.Cells(1, 13).Font.Bold = True
    
    ' 读A列产品编号
    Dim codeArr As Variant
    codeArr = wsRpt.Range("A2:A" & lastRow).Value
    
    Dim nRows As Long: nRows = UBound(codeArr, 1)
    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To 1)
    
    Dim i As Long
    Dim curCode As String
    For i = 1 To nRows
        curCode = Trim(CStr(codeArr(i, 1)))
        If fillDict.Exists(curCode) Then
            outArr(i, 1) = fillDict(curCode)
        Else
            outArr(i, 1) = 0
        End If
    Next i
    
    ' 一次性写入M列
    wsRpt.Range(wsRpt.Cells(2, 13), wsRpt.Cells(1 + nRows, 13)).Value = outArr
    wsRpt.Columns("M").AutoFit
End Sub
