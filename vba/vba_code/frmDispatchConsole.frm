VERSION 5.00
Begin VB.UserForm frmDispatchConsole
   Caption         =   "净值流程控制台"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSteps
      Caption         =   "主流程步骤"
      Height          =   3600
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6900
      Begin VB.ListBox lstSteps
         Height          =   2205
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   6420
      End
      Begin VB.CommandButton cmdRunSelected
         Caption         =   "执行选中步骤"
         Height          =   420
         Left            =   240
         TabIndex        =   2
         Top             =   2820
         Width           =   1980
      End
      Begin VB.CommandButton cmdRunAll
         Caption         =   "按顺序执行全部步骤"
         Height          =   420
         Left            =   2460
         TabIndex        =   3
         Top             =   2820
         Width           =   2220
      End
      Begin VB.CommandButton cmdClose
         Caption         =   "关闭"
         Height          =   420
         Left            =   4920
         TabIndex        =   4
         Top             =   2820
         Width           =   1740
      End
   End
   Begin VB.Frame fraTools
      Caption         =   "工具"
      Height          =   1140
      Left            =   240
      TabIndex        =   5
      Top             =   4020
      Width           =   6900
      Begin VB.CommandButton cmdToolCleanDup
         Caption         =   "清洗重复数据"
         Height          =   420
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   1980
      End
      Begin VB.CommandButton cmdToolPrintUpdate
         Caption         =   "输出更新对象"
         Height          =   420
         Left            =   2460
         TabIndex        =   7
         Top             =   420
         Width           =   1980
      End
      Begin VB.CommandButton cmdToolDelete
         Caption         =   "批量删除数据"
         Height          =   420
         Left            =   4680
         TabIndex        =   8
         Top             =   420
         Width           =   1980
      End
   End
   Begin VB.Label lblStatus
      Caption         =   "请选择步骤后执行，或直接按顺序执行全部步骤。"
      Height          =   660
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   6900
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDispatchConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STEP_COUNT As Long = 6

Private mStepNames(1 To STEP_COUNT) As String

Private Sub UserForm_Initialize()
    mStepNames(1) = "STEP1 导入净值数据"
    mStepNames(2) = "STEP2 生成数据报告"
    mStepNames(3) = "STEP3 导出产品数据"
    mStepNames(4) = "STEP4 生成产品图表"
    mStepNames(5) = "STEP5 导出拼接图表"
    mStepNames(6) = "STEP6 导入全部展示图片"

    Dim i As Long
    For i = 1 To STEP_COUNT
        lstSteps.AddItem mStepNames(i)
    Next i

    lstSteps.ListIndex = 0
    SetStatus "请选择步骤后执行，或直接按顺序执行全部步骤。"
End Sub

Private Sub cmdRunSelected_Click()
    If lstSteps.ListIndex < 0 Then
        MsgBox "请先选择一个要执行的步骤。", vbExclamation, "未选择步骤"
        Exit Sub
    End If

    RunOneStep lstSteps.ListIndex + 1
End Sub

Private Sub cmdRunAll_Click()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("将按 STEP1 到 STEP6 顺序执行全部主流程步骤。" & vbCrLf & _
                    "三个工具按钮不会在此流程中自动执行。" & vbCrLf & vbCrLf & _
                    "是否继续?", _
                    vbQuestion + vbYesNo, "确认执行全部步骤")
    If answer <> vbYes Then Exit Sub

    Dim i As Long
    For i = 1 To STEP_COUNT
        If Not RunOneStep(i) Then Exit For
    Next i
End Sub

Private Sub cmdToolCleanDup_Click()
    RunTool "清洗重复数据", 1
End Sub

Private Sub cmdToolPrintUpdate_Click()
    RunTool "输出更新对象", 2
End Sub

Private Sub cmdToolDelete_Click()
    RunTool "批量删除数据", 3
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Function RunOneStep(ByVal stepIndex As Long) As Boolean
    On Error GoTo ErrHandler

    SetBusy True
    SetStatus "正在执行: " & mStepNames(stepIndex)
    DoEvents

    SaveAndRunStep stepIndex

    SetStatus "已完成: " & mStepNames(stepIndex)
    RunOneStep = True

CleanExit:
    SetBusy False
    DoEvents
    Exit Function

ErrHandler:
    SetStatus "执行失败: " & mStepNames(stepIndex)
    MsgBox "执行失败: " & mStepNames(stepIndex) & vbCrLf & vbCrLf & _
           "错误 " & Err.Number & ": " & Err.Description, _
           vbCritical, "流程中断"
    RunOneStep = False
    Resume CleanExit
End Function

Private Sub RunTool(ByVal toolName As String, ByVal toolIndex As Long)
    On Error GoTo ErrHandler

    SetBusy True
    SetStatus "正在执行工具: " & toolName
    DoEvents

    SaveAndRunTool toolIndex

    SetStatus "工具执行完成: " & toolName

CleanExit:
    SetBusy False
    DoEvents
    Exit Sub

ErrHandler:
    SetStatus "工具执行失败: " & toolName
    MsgBox "工具执行失败: " & toolName & vbCrLf & vbCrLf & _
           "错误 " & Err.Number & ": " & Err.Description, _
           vbCritical, "工具执行失败"
    Resume CleanExit
End Sub

Private Sub SaveAndRunStep(ByVal stepIndex As Long)
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    SaveApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    On Error GoTo CleanFail

    Select Case stepIndex
        Case 1
            STEP1导入净值数据
        Case 2
            STEP2生成数据报告
        Case 3
            STEP3导出产品数据
        Case 4
            STEP4生成产品图表
        Case 5
            STEP5导出拼接图表
        Case 6
            STEP6导入全部展示图片
        Case Else
            Err.Raise vbObjectError + 1001, , "未知步骤编号: " & stepIndex
    End Select

CleanExit:
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    Exit Sub

CleanFail:
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SaveAndRunTool(ByVal toolIndex As Long)
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    SaveApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    On Error GoTo CleanFail

    Select Case toolIndex
        Case 1
            清洗重复数据
        Case 2
            输出更新对象
        Case 3
            批量删除数据
        Case Else
            Err.Raise vbObjectError + 1002, , "未知工具编号: " & toolIndex
    End Select

CleanExit:
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    Exit Sub

CleanFail:
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SaveApplicationState(ByRef oldScreenUpdating As Boolean, _
                                 ByRef oldEnableEvents As Boolean, _
                                 ByRef oldDisplayAlerts As Boolean, _
                                 ByRef oldCalculation As XlCalculation)
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
End Sub

Private Sub RestoreApplicationState(ByVal oldScreenUpdating As Boolean, _
                                    ByVal oldEnableEvents As Boolean, _
                                    ByVal oldDisplayAlerts As Boolean, _
                                    ByVal oldCalculation As XlCalculation)
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
End Sub

Private Sub SetBusy(ByVal isBusy As Boolean)
    cmdRunSelected.Enabled = Not isBusy
    cmdRunAll.Enabled = Not isBusy
    cmdToolCleanDup.Enabled = Not isBusy
    cmdToolPrintUpdate.Enabled = Not isBusy
    cmdToolDelete.Enabled = Not isBusy
    cmdClose.Enabled = Not isBusy
    lstSteps.Enabled = Not isBusy
End Sub

Private Sub SetStatus(ByVal message As String)
    lblStatus.Caption = Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message
End Sub

