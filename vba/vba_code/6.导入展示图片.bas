Attribute VB_Name = "导入展示图片"
Option Explicit

'==============================================================
' 模块: 导入展示图片
' 功能:
'   1. 查找最新的 产品图表_yyyymmdd 文件夹
'   2. 打开 折线图展示\家族家庭-稳享长期限产品收益展示-yyyymmdd.xlsx
'   3. 删除 B4:F4、B5:F5 区域内原有图片
'   4. 插入两张指定红色拼接图
'   5. 根据图片高度调整第4、5行行高
'   6. 另存为更新日期后缀的新文件
'==============================================================

' 如展示文件不是第一个工作表，把这里改成具体sheet名
Private Const DISPLAY_SHEET_NAME As String = ""

Public Sub STEP6导入全部展示图片()

    Dim t0 As Double: t0 = Timer
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim dbPath As String: dbPath = wbDB.Path & "\"
    
    '==========================================================
    ' 自动查找最新图片日期
    '==========================================================
    Dim imgKey As String
    imgKey = FindLatestChartFolderKey(dbPath)
    
    If Len(imgKey) = 0 Then
        MsgBox "未找到[产品图表_yyyymmdd]文件夹,请先运行[导出拼接图表]。", vbExclamation
        GoTo CleanUp
    End If
    
    Dim imgFolder As String
    imgFolder = dbPath & "产品图表_" & imgKey & "\"
    
    If Dir(imgFolder, vbDirectory) = "" Then
        MsgBox "未找到图片文件夹:" & vbCrLf & imgFolder, vbExclamation
        GoTo CleanUp
    End If

    Dim outputFiles As String

    '==========================================================
    ' 1. 家族家庭-稳享长期限产品收益展示
    '==========================================================
    outputFiles = outputFiles & vbCrLf & "  - " & UpdateOneDisplayWorkbook( _
        dbPath:=dbPath, _
        displayPrefix:="家族家庭-稳享长期限产品收益展示-", _
        outputFileName:="家族家庭-稳享长期限产品收益展示-" & imgKey & ".xlsx", _
        imgFolder:=imgFolder, _
        imgNames:=Array( _
            "汇益稳享364天101号_红.png", _
            "汇益稳享728天108号_红.png" _
        ), _
        targetAddresses:=Array( _
            "B4:F4", _
            "B5:F5" _
        ))

    '==========================================================
    ' 2. 交通银行-产品收益展示
    '==========================================================
    outputFiles = outputFiles & vbCrLf & "  - " & UpdateOneDisplayWorkbook( _
        dbPath:=dbPath, _
        displayPrefix:="交通银行-产品收益展示-", _
        outputFileName:="交通银行-产品收益展示-" & imgKey & ".xlsx", _
        imgFolder:=imgFolder, _
        imgNames:=Array( _
            "汇益稳健7天2号_红.png", _
            "汇益稳健28天6号_红.png", _
            "汇益稳享91天3号_红.png" _
        ), _
        targetAddresses:=Array( _
            "B4:F4", _
            "B5:F5", _
            "B6:F6" _
        ))

    '==========================================================
    ' 3. 直销-汇益系列产品收益展示
    '==========================================================
    outputFiles = outputFiles & vbCrLf & "  - " & UpdateOneDisplayWorkbook( _
        dbPath:=dbPath, _
        displayPrefix:="直销-汇益系列产品收益展示-", _
        outputFileName:="直销-汇益系列产品收益展示-" & imgKey & ".xlsx", _
        imgFolder:=imgFolder, _
        imgNames:=Array( _
            "汇益稳健日开101号_红.png", _
            "汇益稳健28天101号_红.png", _
            "交鑫致远6个月101号_红.png" _
        ), _
        targetAddresses:=Array( _
            "B4:F4", _
            "B5:F5", _
            "B6:F6" _
        ))
        
    '==========================================================
    ' 4. 江苏银行-圆融安享产品收益展示
    '==========================================================
    outputFiles = outputFiles & vbCrLf & "  - " & UpdateOneDisplayWorkbook( _
        dbPath:=dbPath, _
        displayPrefix:="江苏银行-圆融安享产品收益展示-", _
        outputFileName:="江苏银行-圆融安享产品收益展示-" & imgKey & ".xlsx", _
        imgFolder:=imgFolder, _
        imgNames:=Array( _
            "raw\圆融安享日开8号_净值_蓝.png", _
            "raw\圆融安享7天2号_净值_蓝.png", _
            "raw\圆融安享28天1号_净值_蓝.png" _
        ), _
        targetAddresses:=Array( _
            "A2:E2", _
            "A13:E13", _
            "A18:E18" _
        ))

    MsgBox "全部展示文件更新完成!" & vbCrLf & _
           "图片日期: " & imgKey & vbCrLf & _
           "处理完成文件:" & outputFiles & vbCrLf & _
           "耗时: " & Format(Timer - t0, "0.00") & " 秒", _
           vbInformation, "处理结果"
    GoTo CleanUp

ErrHandler:
    MsgBox "更新失败:" & vbCrLf & Err.Description, vbCritical, "错误"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

'==============================================================
' 查找最新 产品图表_yyyymmdd 文件夹
'==============================================================
Private Function FindLatestChartFolderKey(ByVal dbPath As String) As String

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^产品图表_(\d{8})$"
    regex.IgnoreCase = True

    Dim folderName As String
    folderName = Dir(dbPath & "产品图表_*", vbDirectory)

    Dim latestKey As String
    Dim matches As Object

    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            If (GetAttr(dbPath & folderName) And vbDirectory) = vbDirectory Then
                If regex.Test(folderName) Then
                    Set matches = regex.Execute(folderName)
                    If matches(0).SubMatches(0) > latestKey Then
                        latestKey = matches(0).SubMatches(0)
                    End If
                End If
            End If
        End If
        folderName = Dir()
    Loop

    FindLatestChartFolderKey = latestKey

End Function

Private Function FindLatestDisplayFileByPrefix(ByVal displayFolderPath As String, _
                                                ByVal displayPrefix As String) As String

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^" & displayPrefix & "(\d{8})\.xlsx$"
    regex.IgnoreCase = True

    Dim fileName As String
    fileName = Dir(displayFolderPath & displayPrefix & "*.xlsx")

    Dim latestKey As String
    Dim latestFile As String
    Dim matches As Object

    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Set matches = regex.Execute(fileName)
            If matches(0).SubMatches(0) > latestKey Then
                latestKey = matches(0).SubMatches(0)
                latestFile = fileName
            End If
        End If
        fileName = Dir()
    Loop

    If Len(latestFile) > 0 Then
        FindLatestDisplayFileByPrefix = displayFolderPath & latestFile
    Else
        FindLatestDisplayFileByPrefix = ""
    End If

End Function

'==============================================================
' 删除与指定区域发生重叠的图片
'==============================================================
Private Sub DeletePicturesInRange(ByVal ws As Worksheet, ByVal targetRng As Range)

    Dim i As Long
    Dim shp As Shape

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)

        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            If ShapeOverlapsRange(shp, targetRng) Then
                shp.Delete
            End If
        End If
    Next i

End Sub

'==============================================================
' 判断Shape是否与Range区域重叠
'==============================================================
Private Function ShapeOverlapsRange(ByVal shp As Shape, ByVal rng As Range) As Boolean

    Dim shpLeft As Double, shpRight As Double
    Dim shpTop As Double, shpBottom As Double

    Dim rngLeft As Double, rngRight As Double
    Dim rngTop As Double, rngBottom As Double

    shpLeft = shp.Left
    shpRight = shp.Left + shp.Width
    shpTop = shp.Top
    shpBottom = shp.Top + shp.Height

    rngLeft = rng.Left
    rngRight = rng.Left + rng.Width
    rngTop = rng.Top
    rngBottom = rng.Top + rng.Height

    ShapeOverlapsRange = Not ( _
        shpRight < rngLeft Or _
        shpLeft > rngRight Or _
        shpBottom < rngTop Or _
        shpTop > rngBottom _
    )

End Function

'==============================================================
' 插入图片:
'   - 左上角对齐目标区域
'   - 锁定纵横比
'   - 宽度适配 B:F 区域宽度
'   - 行高按图片高度自动调整
'==============================================================
Private Sub InsertPictureFitWidthAndSetRowHeight(ByVal ws As Worksheet, _
                                                  ByVal imgPath As String, _
                                                  ByVal targetRng As Range, _
                                                  ByVal rowNum As Long)

    Dim shp As Shape
    Dim picW As Double
    Dim picH As Double
    Dim newRowHeight As Double

    Set shp = ws.Shapes.AddPicture( _
        fileName:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=targetRng.Left, _
        Top:=targetRng.Top, _
        Width:=-1, _
        Height:=-1)

    ' 先按目标宽度等比例缩放，并记录缩放后的真实高度
    With shp
        .Name = "展示图_" & CStr(rowNum)
        .LockAspectRatio = msoTrue
        .Width = targetRng.Width

        picW = .Width
        picH = .Height
    End With

    ' 先设置为只随单元格移动，不随单元格尺寸变化
    shp.Placement = xlMove

    ' 再调整行高。由于不是 xlMoveAndSize，图片不会被拉伸
    newRowHeight = picH + 2

    If newRowHeight > 409.5 Then
        newRowHeight = 409.5
    End If

    ws.Rows(rowNum).RowHeight = newRowHeight

    ' 行高变化后，再重新校准图片位置和尺寸
    With shp
        .LockAspectRatio = msoTrue
        .Left = targetRng.Left
        .Top = targetRng.Top
        .Width = picW
        .Height = picH
        .Placement = xlMove
    End With

End Sub

Private Function UpdateOneDisplayWorkbook(ByVal dbPath As String, _
                                          ByVal displayPrefix As String, _
                                          ByVal outputFileName As String, _
                                          ByVal imgFolder As String, _
                                          ByVal imgNames As Variant, _
                                          ByVal targetAddresses As Variant) As String

    Dim displayFolderPath As String
    displayFolderPath = dbPath & "折线图展示\"

    Dim srcPath As String
    srcPath = FindLatestDisplayFileByPrefix(displayFolderPath, displayPrefix)

    If Len(srcPath) = 0 Or Dir(srcPath) = "" Then
        Err.Raise vbObjectError + 1001, , "未找到展示模板文件: " & displayFolderPath & displayPrefix & "yyyymmdd.xlsx"
    End If

    Dim i As Long
    For i = LBound(imgNames) To UBound(imgNames)
        If Dir(imgFolder & CStr(imgNames(i))) = "" Then
            Err.Raise vbObjectError + 1002, , "未找到图片: " & imgFolder & CStr(imgNames(i))
        End If
    Next i

    If UBound(imgNames) - LBound(imgNames) <> UBound(targetAddresses) - LBound(targetAddresses) Then
        Err.Raise vbObjectError + 1003, , "图片数量与目标区域数量不一致: " & displayPrefix
    End If

    Dim wbShow As Workbook
    Set wbShow = Workbooks.Open(fileName:=srcPath, UpdateLinks:=0)

    ' 默认使用第一个工作表
    Dim wsShow As Worksheet
    Set wsShow = wbShow.Worksheets(1)

    ' 删除旧图片 + 插入新图片
    For i = LBound(imgNames) To UBound(imgNames)

        Dim targetRng As Range
        Set targetRng = wsShow.Range(CStr(targetAddresses(i)))

        DeletePicturesInRange wsShow, targetRng

        InsertPictureFitWidthAndSetRowHeight _
            ws:=wsShow, _
            imgPath:=imgFolder & CStr(imgNames(i)), _
            targetRng:=targetRng, _
            rowNum:=targetRng.row

    Next i

    Dim outPath As String
    outPath = displayFolderPath & outputFileName

    If LCase(outPath) <> LCase(srcPath) Then
        If Dir(outPath) <> "" Then Kill outPath
    End If

    wbShow.SaveAs fileName:=outPath, FileFormat:=xlOpenXMLWorkbook
    wbShow.Close SaveChanges:=False

    UpdateOneDisplayWorkbook = outPath

End Function

