Option Explicit
' HuanweiDZ Macro 版本 1.0.0.4 alpha
' 版本日期：2020-11-09
' 作者：https://github.com/moauris
' 联系方式：mchenf@icloud.com
' 模组：通用工具函数
' 版本更新内容 1.0.0.4
' 整合了所有的工具函数
' 一般只把不对主程序中的全局变量 viewerSheet 进行引用的通用函数放在这里，并且标记为 Public

'函数：从三个单元格中试图生成日期。如果失败那么返回空字符串
Public Function GenerateDateStringFromRange( _
    yearCell As Range, _
    monthCell As Range, _
    dayCell As Range) As String
    
    Dim dateDate As Date
    Dim Year, Mon, Day As String
    Dim OmitDay As Boolean
    OmitDay = True
    Year = Left(yearCell.Value, 4)
    Mon = "-" & monthCell.Value
    Day = ""
    If Len(dayCell.Value) > 0 Then
        Day = "-" & dayCell.Value
        OmitDay = False
    End If
    '试图创建日期对象，如果错误，返回空字符串
    On Error GoTo RETURN_EMPTY
    dateDate = DateValue(Year & Mon & Day)
    On Error GoTo 0
    If OmitDay Then
        GenerateDateStringFromRange = Format(dateDate, "YYYY年MM月")
    Else
        GenerateDateStringFromRange = Format(dateDate, "YYYY年MM月DD日")
    End If
    Exit Function
RETURN_EMPTY:
    GenerateDateStringFromRange = ""
End Function

'试图转换一个单元格的值为double类型。如果失败则返回0
Public Function convertRngDbl(cell As Range) As Double
    On Error GoTo FAIL_CONVERT
    convertRngDbl = CDbl(cell.Value)
    Exit Function
FAIL_CONVERT:
    convertRngDbl = 0
End Function


' 将某个行标记为【已经匹配】
Public Function MarkRowStatus(target As Range, matchStatus As Long)
    Dim InterColor As Long
    Dim fontColor As Long
    Select Case matchStatus
        Case EntryStatus.dzUnmatched
            InterColor = rgbRed
            fontColor = rgbWhite
        Case EntryStatus.dzException
            
        Case EntryStatus.dzPossibleMatch
            InterColor = rgbYellow
            fontColor = rgbBlack
        Case EntryStatus.dzCertain
            InterColor = rgbGreen
            fontColor = rgbWhite
        Case Else
            InterColor = rgbWhite
            fontColor = rgbBlack
    End Select
    With target
        .Interior.Color = InterColor
        .Font.Color = fontColor
    End With
End Function



'在两个区域之间指向箭头, debug用，废止
Function PointArrow(origin As Range, target As Range)
    
    Dim arrowColor1 As Long
    Dim arrowColor2 As Long
    Dim arrowColor3 As Long
    arrowColor1 = RGB(80, 76, 140)
    arrowColor2 = RGB(255, 51, 0)
    arrowColor3 = RGB(80, 104, 53)
    Select Case arrowColor
        Case arrowColor1
            arrowColor = arrowColor2
        Case arrowColor2
            arrowColor = arrowColor3
        Case arrowColor3
            arrowColor = arrowColor1
        Case Else
            arrowColor = arrowColor1
    End Select
    
    Dim startX As Single _
    , startY As Single _
    , endX As Single _
    , endY As Single
    
    startX = origin.Offset(1, 1).Left
    startY = origin.Offset(1, 1).Top
    endX = target.Left
    endY = target.Top
    
    With origin.Borders
        .Weight = 1
        .Color = arrowColor
    End With
    With target.Borders
        .Weight = 1
        .Color = arrowColor
    End With
    
    Dim arrow As Shape
    Set arrow = viewerSheet.Shapes.AddConnector(msoConnectorStraight, startX, startY, endX, endY)
    
    With arrow.Line
        .EndArrowheadStyle = msoArrowheadTriangle
        .Visible = msoTrue
        .ForeColor.RGB = arrowColor
        .Weight = 3
    End With
    
    
    
End Function


'查找到的匹配项目在这里被同步成关系对应, debug用，废止
Sub MakeRelation(CompanyRow As Range, BankRow As Range, _
        intTimesFound As Integer)
    Dim relationSheet As Worksheet
    Dim relationRow As Range
    Set relationSheet = ThisWorkbook.Worksheets("款项关系表")
    Set relationRow = relationSheet.[A1].CurrentRegion.End(xlDown).Offset(1, 0).Resize(1, 5)
    relationRow(1).Value = CompanyRow.Address
    relationRow(2).Formula = CompanyRow.Formula
    relationRow(4).Value = BankRow.Address
    relationRow(3).Formula = BankRow.Formula
    relationRow(5).Value = True
    '如果被找到的次数大于1，那么用颜色标注该行与上行
    If intTimesFound > 1 Then
        relationRow.Offset(-1, 0).Resize(2, 5).Interior.Color = rgbYellow
        relationRow(5).Value = False
        ' 12345
        ' 6789x
        relationRow(5).Offset(-1, 0).Value = False
    End If
    
    Set relationRow = Nothing
    Set relationSheet = Nothing
End Sub

'Open an openfile dialog
Function RunOpenFileDialog() As Worksheet
    Dim fileToOpen
    Dim TargetBook As Workbook
    
    'Open file dialog
    fileToOpen = Application _
        .GetOpenFilename( _
        "97 - 2002 Excel 工作簿 (*.xls), *.xls")
    If fileToOpen <> False Then
        'TODO: 如果有重名的工作簿已经打开，则使用该工作簿
        '似乎是自动的，不需要手动写
        Set TargetBook = Workbooks.Open(fileToOpen)
        Set RunOpenFileDialog = _
            TargetBook.ActiveSheet
        Exit Function
    End If
    Set RunOpenFileDialog = Nothing
    
End Function

Public Function SyncToViewer(inputSheet As Worksheet, _
    outSheet As Worksheet, _
    Side As Long)
    Dim outputRow As Range
    Dim inputRow As Range
    Dim rngYear As Range
    Dim rngMon As Range
    Dim rngDay As Range
    Dim incurDate As String
    Dim nonEmptyCol As Integer



    Set outputRow = Switch( _
        Side = Sides.Company, outSheet.[A3:F3], _
        Side = Sides.Bank, outSheet.[I3:N3])
    ' 123456
    ' ABCDEF
    ' IJKLMN

    '如果侦测到ouputRow有值，那么直接警告并退出
    If outputRow(1).Value <> "" Then
        Dim message As String
        message = "检测到对象显示区域有内容。将不继续执行。请确认"
        MsgBox message, vbCritical, "试图覆盖有内容的显示区域"
        Exit Function
    End If
    Set inputRow = inputSheet.[A6:I6]

    nonEmptyCol = Switch( _
        Side = Sides.Company, 5, _
        Side = Sides.Bank, 4)

    Do While Len(inputRow(nonEmptyCol).Value) > 0
        '生成日期
        Select Case Side
            Case Sides.Company
                Set rngYear = inputRow(1)
                Set rngMon = inputRow(2)
                Set rngDay = inputRow(3)
            Case Sides.Bank
                Set rngYear = inputSheet.[A4]
                Set rngMon = inputRow(1)
                Set rngDay = inputRow(2)
        End Select

        incurDate = GenerateDateStringFromRange _
            (rngYear, rngMon, rngDay)
        ' 如果日期可以形成，则同步其他项目
        If Len(incurDate) > 0 Then
            outputRow(1) = incurDate '发生日期
            Select Case Side
                Case Sides.Company
                    outputRow(2) = inputRow(4) '凭证号
                    outputRow(3) = inputRow(5) '摘要
                    '往来单位跳过
                Case Sides.Bank
                    outputRow(2) = inputRow(3) '凭证号
                    outputRow(3) = inputRow(4) '摘要
                    '往来单位跳过
            End Select
            outputRow(4) = convertRngDbl(inputRow(6)) '借方
            outputRow(5) = convertRngDbl(inputRow(7)) '贷方
            '方向跳过
            outputRow(6) = convertRngDbl(inputRow(9)) '余额
            Set outputRow = outputRow.Offset(1, 0)
        End If
        Set inputRow = inputRow.Offset(1, 0) '进一行
    Loop
End Function

' 给定某个借方或者贷方区域，在对面的相应位置列找出符合要求的查找区域
' 【符合要求的区域】：1、正负号一致；2、绝对值小于目标绝对值
Public Function SeekConsolidationRange(sourceRange As Range) As String
    Dim sourceSum As Double
    sourceSum = WorksheetFunction.Sum(sourceRange)
    
    Dim targetRange As Range '代表了对面的相应位置
    Dim resultRange As Range
    'D => L, E => M
    '4 ={+8}=> 12, 5 => 13
    Select Case sourceRange.Column
        Case 4
            Set targetRange = _
            sourceRange.Worksheet.[L3]
        Case 5
            Set targetRange = _
            sourceRange.Worksheet.[M3]
        Case 12
            Set targetRange = _
            sourceRange.Worksheet.[D3]
        Case 13
            Set targetRange = _
            sourceRange.Worksheet.[E3]
    End Select
    
    If sourceSum < 0 Then
        Do While targetRange.Value <> ""
            If targetRange.Value < 0 And _
                targetRange.Value > sourceSum Then
                If resultRange Is Nothing Then
                    Set resultRange = targetRange
                Else
                    Set resultRange = Union(resultRange, targetRange)
                End If
                
            End If
            Set targetRange = targetRange.Offset(1, 0)
        Loop
    Else
        Do While targetRange.Value <> ""
            If targetRange.Value > 0 And _
                targetRange.Value < sourceSum Then
                If resultRange Is Nothing Then
                    Set resultRange = targetRange
                Else
                    Set resultRange = Union(resultRange, targetRange)
                End If
                
            End If
            Set targetRange = targetRange.Offset(1, 0)
        Loop
    End If
    SeekConsolidationRange = resultRange.Address
End Function

' 返回某一格当前所在的区域中的整行交际
Public Function SelectCurrentRegionRow(sourceRange As Range) As Range
    Set SelectCurrentRegionRow = _
        Intersect(sourceRange.EntireRow, sourceRange.CurrentRegion)

End Function



