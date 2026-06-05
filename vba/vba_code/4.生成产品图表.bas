Attribute VB_Name = "生成产品图表"
Option Explicit

'==============================================================
' 模块: 生成产品图表
' 功能: 在最新的[产品净值汇总_yyyymmdd.xlsx]中,
'       为每个产品sheet生成4个chart(净值/收益率 × 红/蓝)
'       - 收益率chart跳过开头连续的0值
'       - X轴强制为日期轴,只显示首尾日期
'==============================================================

' 模板文件名(与.xlsm同目录)
Private Const TPL_NAV_RED As String = "净值图表_红.crtx"
Private Const TPL_NAV_BLUE As String = "净值图表_蓝.crtx"
Private Const TPL_YIELD_RED As String = "收益率图表_红.crtx"
Private Const TPL_YIELD_BLUE As String = "收益率图表_蓝.crtx"

' chart布局参数(单位:磅)
Private Const CHART_LEFT_COL As String = "L"
Private Const CHART_WIDTH As Single = 480
Private Const CHART_HEIGHT As Single = 280
Private Const CHART_GAP As Single = 20

' 用于收集"触发Y轴下限0限制"的产品
Private g_clipProducts As Object

Public Sub STEP4生成产品图表()
    
    Dim t0 As Double: t0 = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim dbPath As String: dbPath = wbDB.Path & "\"
    
    '--- 1. 检查4个模板文件 ---
    Dim tplPaths(1 To 4) As String
    tplPaths(1) = dbPath & TPL_NAV_RED
    tplPaths(2) = dbPath & TPL_NAV_BLUE
    tplPaths(3) = dbPath & TPL_YIELD_RED
    tplPaths(4) = dbPath & TPL_YIELD_BLUE
    
    Dim i As Long, missing As String
    For i = 1 To 4
        If Dir(tplPaths(i)) = "" Then
            missing = missing & vbCrLf & "  - " & Mid(tplPaths(i), InStrRev(tplPaths(i), "\") + 1)
        End If
    Next i
    If Len(missing) > 0 Then
        MsgBox "缺少模板文件,请检查:" & missing, vbCritical
        GoTo CleanUp
    End If
    
    '--- 2. 查找最新的[产品净值汇总_yyyymmdd.xlsx] ---
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^产品净值汇总_(\d{8})\.xlsx$"
    
    Dim fileName As String, latestFile As String, latestKey As String
    fileName = Dir(dbPath & "产品净值汇总_*.xlsx")
    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Dim matches As Object
            Set matches = regex.Execute(fileName)
            If matches(0).SubMatches(0) > latestKey Then
                latestKey = matches(0).SubMatches(0)
                latestFile = fileName
            End If
        End If
        fileName = Dir()
    Loop
    
    If Len(latestFile) = 0 Then
        MsgBox "未找到[产品净值汇总_yyyymmdd.xlsx]文件,请先运行[导出产品数据]。", vbExclamation
        GoTo CleanUp
    End If
    
    Dim targetPath As String: targetPath = dbPath & latestFile
    
    '--- 3. 检查目标文件是否已打开 ---
    Dim wbTarget As Workbook
    Dim wasOpen As Boolean: wasOpen = False
    
    On Error Resume Next
    Set wbTarget = Workbooks(latestFile)
    On Error GoTo 0
    
    If wbTarget Is Nothing Then
        Set wbTarget = Workbooks.Open(targetPath)
    Else
        wasOpen = True
    End If
    
    '--- 4. 遍历每个sheet,生成4个chart ---
    Dim ws As Worksheet
    Dim processedCount As Long: processedCount = 0
    Dim errSheets As String: errSheets = ""
    
    ' 初始化收集器
    Set g_clipProducts = CreateObject("Scripting.Dictionary")
    
    For Each ws In wbTarget.Worksheets
        ' 检查sheet是否有有效数据
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If lastRow < 2 Then GoTo NextSheet
        
        ' 先清除该sheet上已有的chart
        Dim co As ChartObject
        For Each co In ws.ChartObjects
            co.Delete
        Next co
        
        ' 净值的数据范围: 从第2行(首个数据行)到末行
        Dim navStartRow As Long: navStartRow = 2
        Dim dateRngNav As Range, navRng As Range
        Set dateRngNav = ws.Range("A" & navStartRow & ":A" & lastRow)
        Set navRng = ws.Range("D" & navStartRow & ":D" & lastRow)
        
        ' 收益率的数据范围: 跳过开头连续的0
        Dim yieldStartRow As Long
        yieldStartRow = FindFirstNonZeroRow(ws, "G", 2, lastRow)
        
        Dim dateRngYield As Range, yieldRng As Range
        Dim hasYieldData As Boolean: hasYieldData = False
        If yieldStartRow > 0 And yieldStartRow <= lastRow Then
            Set dateRngYield = ws.Range("A" & yieldStartRow & ":A" & lastRow)
            Set yieldRng = ws.Range("G" & yieldStartRow & ":G" & lastRow)
            hasYieldData = True
        End If
        
        Dim prodShort As String: prodShort = ws.Name
        
        ' 计算4个chart的位置(2列2行)
        Dim baseLeft As Single, baseTop As Single
        baseLeft = ws.Range(CHART_LEFT_COL & "1").Left
        baseTop = ws.Range(CHART_LEFT_COL & "1").Top
        
        On Error GoTo SheetErr
        
        ' Chart 1: 净值-红 (左上)
        CreateChart ws, dateRngNav, navRng, tplPaths(1), _
            "chart_净值_红", prodShort & "成立以来净值表现", _
            baseLeft, baseTop
        
        ' Chart 2: 净值-蓝 (左下) ← 改这里
        CreateChart ws, dateRngNav, navRng, tplPaths(2), _
            "chart_净值_蓝", prodShort & "成立以来净值表现", _
            baseLeft, baseTop + CHART_HEIGHT + CHART_GAP
        
        ' Chart 3 & 4: 收益率(强制Y轴下限为0)
        If hasYieldData Then
            ' Chart 3: 收益率-红 (右上)
            CreateChart ws, dateRngYield, yieldRng, tplPaths(3), _
                "chart_收益率_红", prodShort & "成立以来30日年化收益率" & vbLf & "(单位:%)", _
                baseLeft + CHART_WIDTH + CHART_GAP, baseTop, True
            
            ' Chart 4: 收益率-蓝 (右下)
            CreateChart ws, dateRngYield, yieldRng, tplPaths(4), _
                "chart_收益率_蓝", prodShort & "成立以来30日年化收益率" & vbLf & "(单位:%)", _
                baseLeft + CHART_WIDTH + CHART_GAP, baseTop + CHART_HEIGHT + CHART_GAP, True
        End If
        
        processedCount = processedCount + 1
        On Error GoTo 0
        GoTo NextSheet

SheetErr:
        errSheets = errSheets & ws.Name & "(" & Err.Description & "), "
        Err.Clear
        On Error GoTo 0

NextSheet:
    Next ws
    
    '--- 5. 保存文件 ---
    wbTarget.Save
    If Not wasOpen Then wbTarget.Close SaveChanges:=False
    
    '--- 6. 汇总提示 ---
    Dim msg As String
    msg = "图表生成完成!" & vbCrLf & _
          "目标文件: " & latestFile & vbCrLf & _
          "处理sheet数: " & processedCount & vbCrLf & _
          "耗时: " & Format(Timer - t0, "0.00") & " 秒"
    If Len(errSheets) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "以下sheet处理出错:" & vbCrLf & Left(errSheets, Len(errSheets) - 2)
    End If
    
    ' 列出触发"收益率Y轴下限强制为0"的产品
    If g_clipProducts.Count > 0 Then
        Dim clipList As String: clipList = ""
        Dim k As Variant
        For Each k In g_clipProducts.keys
            clipList = clipList & CStr(k) & ", "
        Next k
        msg = msg & vbCrLf & vbCrLf & _
              "以下产品的收益率Y轴下限被强制设为0(原计算值为负):" & vbCrLf & _
              Left(clipList, Len(clipList) - 2)
    End If
    
    MsgBox msg, vbInformation, "处理结果"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


'==============================================================
' 辅助函数: 找到指定列从startRow到endRow中,第一个非0(且非空)的行号
' 如果全部为0或空,返回0
'==============================================================
Private Function FindFirstNonZeroRow(ByVal ws As Worksheet, ByVal col As String, _
                                      ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    Dim v As Variant
    For r = startRow To endRow
        v = ws.Range(col & r).Value
        If Not IsEmpty(v) And IsNumeric(v) Then
            If CDbl(v) <> 0 Then
                FindFirstNonZeroRow = r
                Exit Function
            End If
        End If
    Next r
    FindFirstNonZeroRow = 0  ' 全0或全空
End Function

'==============================================================
' 辅助过程: 创建一个chart并应用模板
'==============================================================
Private Sub CreateChart(ByVal ws As Worksheet, _
                       ByVal xRange As Range, ByVal yRange As Range, _
                       ByVal tplPath As String, _
                       ByVal chartName As String, ByVal title As String, _
                       ByVal leftPos As Single, ByVal topPos As Single, _
                       Optional ByVal enforceYMinZero As Boolean = False)
    
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=leftPos, Top:=topPos, _
                                  Width:=CHART_WIDTH, Height:=CHART_HEIGHT)
    co.Name = chartName
    
    Dim ch As Chart
    Set ch = co.Chart
    
    ' 清空默认系列
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    
    ' 添加系列
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Values = yRange
    s.XValues = xRange
    s.Name = title
    
    ' 应用模板
    On Error Resume Next
    ch.ApplyChartTemplate tplPath
    On Error GoTo 0
    
    ' 重新拿系列引用(ApplyChartTemplate可能让原引用失效)
    Set s = ch.SeriesCollection(1)
    
    ' 设置标题
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    
    ' 标题文字右对齐(短行向右靠拢,长行视觉上保持居中)
    On Error Resume Next
    ch.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
    On Error GoTo 0
    
    ' --- 强制X轴为日期轴,只显示首尾日期 ---
    Dim firstDate As Date, lastDate As Date
    Dim dayDiff As Long
    
    On Error Resume Next
    firstDate = CDate(xRange.Cells(1, 1).Value)
    lastDate = CDate(xRange.Cells(xRange.Rows.Count, 1).Value)
    On Error GoTo 0
    
    If firstDate > 0 And lastDate > firstDate Then
        dayDiff = lastDate - firstDate
        
        On Error Resume Next
        With ch.Axes(xlCategory)
            .CategoryType = xlTimeScale
            .MinimumScale = CDbl(firstDate)
            .MaximumScale = CDbl(lastDate)
            .MajorUnit = dayDiff
            .MajorUnitScale = xlDays
            .MinorUnit = dayDiff
            .MinorUnitScale = xlDays
            .TickLabels.NumberFormat = "yyyy-mm-dd"
        End With
        On Error GoTo 0
    End If
    
    ' --- Y轴自适应 ---
    Dim yMin As Double, yMax As Double
    Dim hasValid As Boolean
    hasValid = GetMinMaxFromRange(yRange, yMin, yMax)
    
    If hasValid Then
        If yMax > yMin Then
            Dim padding As Double
            padding = (yMax - yMin) * 0.1
            
            Dim finalMin As Double, finalMax As Double
            finalMin = yMin - padding
            finalMax = yMax + padding
            
            ' --- 对齐到档位 ---
            ' enforceYMinZero为True表示这是收益率chart,档位0.1; 否则净值chart,档位0.001
            Dim stepUnit As Double
            If enforceYMinZero Then
                stepUnit = 0.1      ' 收益率: 1位小数
            Else
                stepUnit = 0.001    ' 净值: 3位小数
            End If
            
            finalMin = AlignDown(finalMin, stepUnit)
            finalMax = AlignUp(finalMax, stepUnit)
            
            ' --- 如果启用"下限强制为0",且对齐后下限仍<0,则截断为0 ---
            If enforceYMinZero And finalMin < 0 Then
                finalMin = 0
                ' 记录该产品触发了限制
                If Not g_clipProducts Is Nothing Then
                    g_clipProducts(ws.Name) = 1
                End If
            End If
            
            On Error Resume Next
            With ch.Axes(xlValue)
                .MinimumScale = finalMin
                .MaximumScale = finalMax
            End With
            On Error GoTo 0
        End If
    End If
    
    ' --- 数据标签: 先全开,再删除非末尾的 ---
    ' 这种做法在某些Excel版本下比直接Point(N).HasDataLabel=True更可靠
    s.HasDataLabels = True
    
    Dim ptCount As Long
    ptCount = s.Points.Count
    
    Dim p As Long
    For p = 1 To ptCount
        If p < ptCount Then
            ' 非末尾点,删除标签
            s.Points(p).HasDataLabel = False
        Else
            ' 末尾点,配置标签
            With s.Points(p).DataLabel
                .ShowValue = True
                .ShowCategoryName = False
                .ShowSeriesName = False
                .NumberFormat = "0.0000"
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .Font.Bold = True
            End With
        End If
    Next p
    
    ' --- 强制ChartArea背景为白色,无边框 ---
    On Error Resume Next
    With ch.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
    With ch.ChartArea.Format.Line
        .Visible = msoFalse
    End With
    On Error GoTo 0
    
    ' --- 字体设置: 西文Times New Roman, 中文仿宋, 复杂文字Times New Roman ---
    On Error Resume Next
    ' 批量设置全图字体
    With ch.ChartArea.Format.TextFrame2.TextRange.Font
        .Name = "Times New Roman"
        .NameFarEast = "仿宋"
        .NameComplexScript = "Times New Roman"
    End With
    
    ' 标题字体单独设置
    If ch.HasTitle Then
        With ch.ChartTitle.Format.TextFrame2.TextRange.Font
            .Name = "Times New Roman"
            .NameFarEast = "仿宋"
            .NameComplexScript = "Times New Roman"
        End With
    End If
    
    ' X轴/Y轴刻度标签字体
    With ch.Axes(xlCategory).TickLabels.Font
        .Name = "Times New Roman"
    End With
    With ch.Axes(xlValue).TickLabels.Font
        .Name = "Times New Roman"
    End With
    On Error GoTo 0
End Sub
'==============================================================
' 辅助函数: 从Range中获取数值的最小值和最大值
' 忽略空单元格和非数值,返回是否找到至少一个有效值
'==============================================================
Private Function GetMinMaxFromRange(ByVal rng As Range, _
                                     ByRef outMin As Double, ByRef outMax As Double) As Boolean
    Dim arr As Variant
    Dim isSingleCell As Boolean
    
    ' 单格Range读出来不是数组,需要兼容
    If rng.Cells.Count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = rng.Value
    Else
        arr = rng.Value
    End If
    
    Dim r As Long, c As Long
    Dim v As Variant
    Dim found As Boolean: found = False
    Dim minVal As Double, maxVal As Double
    
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            v = arr(r, c)
            If Not IsEmpty(v) And IsNumeric(v) Then
                Dim d As Double: d = CDbl(v)
                If Not found Then
                    minVal = d
                    maxVal = d
                    found = True
                Else
                    If d < minVal Then minVal = d
                    If d > maxVal Then maxVal = d
                End If
            End If
        Next c
    Next r
    
    If found Then
        outMin = minVal
        outMax = maxVal
    End If
    GetMinMaxFromRange = found
End Function
'==============================================================
' 辅助函数: 向下取整到指定档位(stepUnit的整数倍)
'   例: AlignDown(1.0014, 0.001) = 1.001
'       AlignDown(2.34, 0.1)     = 2.3
'       AlignDown(-0.05, 0.1)    = -0.1  (负数也向下=更负)
'==============================================================
Private Function AlignDown(ByVal v As Double, ByVal stepUnit As Double) As Double
    If stepUnit <= 0 Then
        AlignDown = v
        Exit Function
    End If
    AlignDown = Int(v / stepUnit) * stepUnit
End Function

'==============================================================
' 辅助函数: 向上取整到指定档位
'   例: AlignUp(1.0014, 0.001) = 1.002
'       AlignUp(2.34, 0.1)     = 2.4
'       AlignUp(2.30, 0.1)     = 2.3  (已对齐则不变)
'==============================================================
Private Function AlignUp(ByVal v As Double, ByVal stepUnit As Double) As Double
    If stepUnit <= 0 Then
        AlignUp = v
        Exit Function
    End If
    
    Dim q As Double
    q = v / stepUnit
    
    ' 用一个小容差判断"是否已经是整数倍",避免浮点误差导致已对齐的值被多推一档
    Const EPS As Double = 0.000000001
    If Abs(q - Int(q)) < EPS Then
        ' 已对齐,不动
        AlignUp = Int(q) * stepUnit
    Else
        AlignUp = (Int(q) + 1) * stepUnit
    End If
End Function

