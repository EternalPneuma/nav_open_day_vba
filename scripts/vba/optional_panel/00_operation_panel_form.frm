Option Explicit

' UserForm 名称：frmOperationPanel
' 用途：为 chart、data、weekly_recommendation、product_one_page 与 tool 提供简单操作面板。
' 说明：本窗体会在 Initialize 中动态生成控件，不需要手工绘制按钮。

Private mButtonHandlers As Collection
Private mStatusLabel As MSForms.Label
Private mChartFrame As MSForms.Frame
Private mDataFrame As MSForms.Frame
Private mRecommendationFrame As MSForms.Frame
Private mOnePageFrame As MSForms.Frame
Private mToolFrame As MSForms.Frame

Private Sub UserForm_Initialize()
    Me.Caption = "上层产品净值自动化操作面板"
    Me.Width = 1265
    Me.Height = 455
    Me.BackColor = RGB(248, 248, 248)

    Set mButtonHandlers = New Collection

    AddTitleLabel "titleMain", "自动化操作面板", 18, 14, 1225, 24, 12, True
    AddTitleLabel "titleDesc", "请按业务流程顺序运行；一键流程会依次调用下方各步骤。", 18, 42, 1225, 18, 9, False

    Set mChartFrame = AddSection("frameChart", "图表流程 chart", 18, 76, 225, 255)
    AddActionButton mChartFrame, "btnChartAll", "一键运行图表流程", 16, 28, 190, 30, Array("Chart01_ImportNavData", "Chart02_ExportProductSummary", "Chart03_GenerateCharts", "Chart04_ExportImages")
    AddActionButton mChartFrame, "btnChartStep1", "1. 导入净值数据(净值)", 16, 70, 190, 28, Array("Chart01_ImportNavData")
    AddActionButton mChartFrame, "btnChartStep2", "2. 输出产品净值汇总", 16, 106, 190, 28, Array("Chart02_ExportProductSummary")
    AddActionButton mChartFrame, "btnChartStep3", "3. 生成产品图表", 16, 142, 190, 28, Array("Chart03_GenerateCharts")
    AddActionButton mChartFrame, "btnChartStep4", "4. 导出产品图片", 16, 178, 190, 28, Array("Chart04_ExportImages")

    Set mDataFrame = AddSection("frameData", "数据流程 data", 263, 76, 225, 255)
    AddActionButton mDataFrame, "btnDataAll", "一键运行数据流程", 16, 28, 190, 30, Array("Data01_ImportNav181", "Data02_CalculateOpenDate", "Data03_ExportProductReport", "Data04_ExportDisplayReport")
    AddActionButton mDataFrame, "btnDataStep1", "1. 导入净值数据(181)", 16, 70, 190, 28, Array("Data01_ImportNav181")
    AddActionButton mDataFrame, "btnDataStep2", "2. 测算开放日", 16, 106, 190, 28, Array("Data02_CalculateOpenDate")
    AddActionButton mDataFrame, "btnDataStep3", "3. 输出分类表现", 16, 142, 190, 28, Array("Data03_ExportProductReport")
    AddActionButton mDataFrame, "btnDataStep4", "4. 输出展示报表", 16, 178, 190, 28, Array("Data04_ExportDisplayReport")

    Set mRecommendationFrame = AddSection("frameRecommendation", "推荐材料 weekly_recommendation", 508, 76, 225, 255)
    AddActionButton mRecommendationFrame, "btnRecommendationAll", "一键运行推荐材料流程", 16, 28, 190, 30, Array("Weekly01_UpdateDependencies", "Weekly02_GenerateReport")
    AddActionButton mRecommendationFrame, "btnRecommendationStep1", "1. 更新推荐材料依赖", 16, 70, 190, 28, Array("Weekly01_UpdateDependencies")
    AddActionButton mRecommendationFrame, "btnRecommendationStep2", "2. 生成推荐材料", 16, 106, 190, 28, Array("Weekly02_GenerateReport")

    Set mOnePageFrame = AddSection("frameOnePage", "单产品一页 product_one_page", 753, 76, 225, 255)
    AddActionButton mOnePageFrame, "btnOnePageAll", "一键运行单产品一页流程", 16, 28, 190, 30, Array("OnePage00_CheckAndImportNavData", "OnePage01_ExportChartData", "OnePage02_GenerateCharts", "OnePage03_ExportPptPdf")
    AddActionButton mOnePageFrame, "btnOnePageStep0", "0. 检查并补充净值数据", 16, 70, 190, 28, Array("OnePage00_CheckAndImportNavData")
    AddActionButton mOnePageFrame, "btnOnePageStep1", "1. 导出一页通数据", 16, 106, 190, 28, Array("OnePage01_ExportChartData")
    AddActionButton mOnePageFrame, "btnOnePageStep2", "2. 生成一页通图表", 16, 142, 190, 28, Array("OnePage02_GenerateCharts")
    AddActionButton mOnePageFrame, "btnOnePageStep3", "3. 导出 PPT/PDF", 16, 178, 190, 28, Array("OnePage03_ExportPptPdf")

    Set mToolFrame = AddSection("frameTool", "维护工具 tool", 998, 76, 225, 255)
    AddActionButton mToolFrame, "btnToolCleanData", "1. 【绘图净值数据】去重", 16, 70, 190, 28, Array("Tool01_CleanDuplicateData")
    AddActionButton mToolFrame, "btnToolDeleteData", "2. 按产品编号删除数据", 16, 106, 190, 28, Array("Tool02_DeleteByProductId")
    AddActionButton mToolFrame, "btnToolFillOpenDate", "3. 补充下一开放日", 16, 142, 190, 28, Array("Tool03_FillNextOpenDate")
    AddActionButton mToolFrame, "btnToolCheckNavData", "4. 核对净值数据", 16, 178, 190, 28, Array("Tool04_CheckNavData")

    AddActionButton Me, "btnClose", "关闭面板", 1130, 348, 96, 30, Array("__close__")

    Set mStatusLabel = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With mStatusLabel
        .Left = 18
        .Top = 348
        .Width = 1085
        .Height = 54
        .Caption = "状态：等待操作。"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .WordWrap = True
    End With
End Sub

Public Sub RunPanelAction(ByVal actionTitle As String, ByVal macroNames As Variant)
    If IsArray(macroNames) Then
        If UBound(macroNames) >= LBound(macroNames) Then
            If CStr(macroNames(LBound(macroNames))) = "__close__" Then
                Unload Me
                Exit Sub
            End If
        End If
    End If

    Dim confirmText As String
    confirmText = "确认运行：" & actionTitle & "？" & vbCrLf & vbCrLf & _
                  "运行前请确认当前工作簿已保存，且相关源文件已经放在同级目录。"
    If MsgBox(confirmText, vbQuestion + vbYesNo + vbDefaultButton2, "确认运行") <> vbYes Then Exit Sub

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    On Error GoTo RunFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim i As Long
    For i = LBound(macroNames) To UBound(macroNames)
        SetStatus "正在运行：" & CStr(macroNames(i))
        DoEvents
        RunMacroByName CStr(macroNames(i))
    Next i

    SetStatus "完成：" & actionTitle

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Exit Sub

RunFail:
    SetStatus "失败：" & actionTitle & vbCrLf & Err.Description
    MsgBox "运行失败：" & actionTitle & vbCrLf & vbCrLf & Err.Description, vbExclamation, "操作面板"
    Resume CleanExit
End Sub

Private Sub RunMacroByName(ByVal macroName As String)
    ' VBE 中所有模块通常都在同一层级；这里按过程名直接调用即可。
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
End Sub

Private Sub SetStatus(ByVal statusText As String)
    If Not mStatusLabel Is Nothing Then
        mStatusLabel.Caption = "状态：" & statusText
    End If
End Sub

Private Function AddSection(ByVal controlName As String, _
                            ByVal captionText As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal controlWidth As Single, _
                            ByVal controlHeight As Single) As MSForms.Frame
    Dim frameCtl As MSForms.Frame
    Set frameCtl = Me.Controls.Add("Forms.Frame.1", controlName, True)
    With frameCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .ForeColor = RGB(50, 50, 50)
    End With
    Set AddSection = frameCtl
End Function

Private Sub AddTitleLabel(ByVal controlName As String, _
                          ByVal captionText As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal controlWidth As Single, _
                          ByVal controlHeight As Single, _
                          ByVal fontSize As Single, _
                          ByVal isBold As Boolean)
    Dim labelCtl As MSForms.Label
    Set labelCtl = Me.Controls.Add("Forms.Label.1", controlName, True)
    With labelCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(35, 35, 35)
        .Font.Name = "微软雅黑"
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
End Sub

Private Sub AddActionButton(ByVal parentControl As Object, _
                            ByVal controlName As String, _
                            ByVal captionText As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal controlWidth As Single, _
                            ByVal controlHeight As Single, _
                            ByVal macroNames As Variant)
    Dim buttonCtl As MSForms.CommandButton
    Set buttonCtl = parentControl.Controls.Add("Forms.CommandButton.1", controlName, True)
    With buttonCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .TakeFocusOnClick = False
    End With

    Dim buttonHandler As clsOperationPanelButton
    Set buttonHandler = New clsOperationPanelButton
    buttonHandler.Init buttonCtl, Me, captionText, macroNames
    mButtonHandlers.Add buttonHandler
End Sub
