Option Explicit
' HuanweiDZ Macro 版本 1.0.0.1 alpha
' 版本日期：2020-11-06
' 作者：https://github.com/moauris
' 联系方式：mchenf@icloud.com
Dim viewerSheet As Worksheet '代表了表格显示区
Dim coRegion, baRegion As Range '代表了公司、银行方工作区域
Dim arrowColor As Long '全局变量箭头颜色
Dim i As Integer
Dim j As Integer
Dim rng As Range

Dim emptyFormula(1 To 7) As Variant '全局变量，用于表示一行空值
' Enumeration for the Status of the Entries
Public Enum EntryStatus
    dzUnmatched = 0
    dzException = 1
    dzPossibleMatch = 2
    dzCertain = 3
End Enum

Sub SyncFromBook_btn_Click()
    Call SyncFromBookMain
End Sub

Sub ClearAll_btn_Click()
    Set viewerSheet = ThisWorkbook.Worksheets("表格显示区")
    Dim clearArea As Range
    Dim shp As Shape
    Set clearArea = viewerSheet.[A1].CurrentRegion
    clearArea.Offset(2, 0).Clear
    Set clearArea = viewerSheet.[I1].CurrentRegion
    clearArea.Offset(2, 0).Clear
    viewerSheet.[H:H].Clear
    For Each shp In viewerSheet.Shapes
        shp.Delete
    Next shp
End Sub

'Main Procedure for Sync from book
Sub SyncFromBookMain()
    Dim targetSheet As Worksheet
    Dim LedgerTitle As String
    'Run OpenFile Dialog
    Set targetSheet = RunOpenFileDialog
    Set viewerSheet = ThisWorkbook.Worksheets("表格显示区")
    ' 初始化空的方程列（7）
    emptyFormula(1) = "'-"
    emptyFormula(2) = "'-"
    emptyFormula(3) = "'-"
    emptyFormula(4) = 0
    emptyFormula(5) = 0
    emptyFormula(6) = 0
    emptyFormula(7) = 0
    If targetSheet Is Nothing Then Exit Sub
    LedgerTitle = targetSheet.[A1].Value
    
    'Debug.Print (LedgerTitle)
    '辅助明细账  = Bank
    '科目明细账 = Company
    Select Case Left(LedgerTitle, 5)
        Case "辅助明细账"
            Call BankSyncTo(targetSheet)
        Case "科目明细账"
            Call CompanySyncTo(targetSheet)
        Case Else
            '没有找到合适的抬头文字，直接退出
            Dim message As String
            message = "请检查导入的表格是否是一个合适的文件，或者抬头是否为“辅助明细账”或者“科目明细账”，并且不含有特殊符号"
            MsgBox message, vbCritical, "表格不符合导入规范"
            GoTo CLEAN_UP
    End Select
    '导入和单项对账完成，标记双方工作区
    Dim CurrencyColumns As Range
    '对账完成，开始制造中文货币格式
    Set CurrencyColumns = viewerSheet.[D:G,L:O]
    CurrencyColumns.NumberFormat = _
        "_ [$￥-zh-CN]* #,##0.00_ ;_ [$￥-zh-CN]* -#,##0.00_ ;_ [$￥-zh-CN]* ""-""??_ ;_ @_ "
CLEAN_UP:
    Set targetSheet = Nothing
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
        Set TargetBook = Workbooks.Open(fileToOpen)
        Set RunOpenFileDialog = _
            TargetBook.ActiveSheet
        Exit Function
    End If
    Set RunOpenFileDialog = Nothing
    
    
End Function
'同步到银行区域
Sub BankSyncTo(inputSheet As Worksheet)
    '定义区域
    Dim targetSheet As Worksheet
    Dim targetRow, inputRow As Range
    Dim incurredDate As String
    
    Set targetRow = viewerSheet.[I3:O3]
    'I J K L M N O
    '1 2 3 4 5 6 7
    Set inputRow = inputSheet.[A6:I6]
    'ABCDEFGHI
    '123456789A
    'incurredDate = GenerateDateStringFromRange _
        (inputSheet.[A4], inputRow(1), inputRow(2))
        
    'Debug.Print (incurredDate)
    
    '循环每一行inputSheet，第4列不为空时循环
    Do While Len(inputRow(4).Value) > 0
        '生成日期
        incurredDate = GenerateDateStringFromRange _
            (inputSheet.[A4], inputRow(1), inputRow(2))
        If Len(incurredDate) > 0 Then
            targetRow(1) = incurredDate '发生日期
            targetRow(2) = inputRow(3) '凭证号
            targetRow(3) = inputRow(4) '摘要
            '往来单位跳过
            
            targetRow(4) = convertRngDbl(inputRow(6)) '借方
            
            targetRow(5) = convertRngDbl(inputRow(7)) '贷方
            '方向跳过
            '贷方余额计算
            targetRow(6) = targetRow(4) - targetRow(5)
            '如果为负数那么字体变红
            If targetRow(6).Value < 0 Then targetRow(6).Font.Color = rgbRed
            targetRow(7) = convertRngDbl(inputRow(9)) '余额
            Set targetRow = targetRow.Offset(1, 0)
        End If
        Set inputRow = inputRow.Offset(1, 0) '进一行
    Loop


    Call TryConsolidateSingle
End Sub
''同步到公司区域
Sub CompanySyncTo(inputSheet As Worksheet)
    '定义区域
    Dim targetRow, inputRow As Range
    Dim incurredDate As String
    
    Set targetRow = viewerSheet.[A3:G3]
    'I J K L M N O
    '1 2 3 4 5 6 7
    Set inputRow = inputSheet.[A6:I6]
    'ABCDEFGHI
    '123456789A
    'incurredDate = GenerateDateStringFromRange _
        (inputSheet.[A4], inputRow(1), inputRow(2))
        
    'Debug.Print (incurredDate)
    
    '循环每一行inputSheet，第5列不为空时循环
    Do While Len(inputRow(5).Value) > 0
        '生成日期
        incurredDate = GenerateDateStringFromRange _
            (inputRow(1), inputRow(2), inputRow(3))
        If Len(incurredDate) > 0 Then
            targetRow(1) = incurredDate '发生日期
            targetRow(2) = inputRow(4) '凭证号
            targetRow(3) = inputRow(5) '摘要
            '往来单位跳过
            
            targetRow(4) = convertRngDbl(inputRow(6)) '借方
            
            targetRow(5) = convertRngDbl(inputRow(7)) '贷方
            '方向跳过
            '贷方余额计算
            targetRow(6) = targetRow(4) - targetRow(5)
            '如果为负数那么字体变红
            If targetRow(6).Value < 0 Then targetRow(6).Font.Color = rgbRed
            targetRow(7) = convertRngDbl(inputRow(9)) '余额
            Set targetRow = targetRow.Offset(1, 0)
        End If
        Set inputRow = inputRow.Offset(1, 0) '进一行
    Loop

    Call TryConsolidateSingle
End Sub
'函数：从三个单元格中试图生成日期。如果失败那么返回空字符串
Function GenerateDateStringFromRange( _
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
Function convertRngDbl(cell As Range) As Double
    On Error GoTo FAIL_CONVERT
    convertRngDbl = CDbl(cell.Value)
    Exit Function
FAIL_CONVERT:
    convertRngDbl = 0
End Function
'检查两表是否齐全，如果齐全，开始对账程序
Sub TryConsolidateSingle()
    '检查两表是否齐全
    Dim CompanyRegion, BankRegion As Range
    Dim colorConsolidated As Variant
    colorConsolidated = RGB(102, 240, 255) '淡蓝色
    Set coRegion = viewerSheet.[A1].CurrentRegion
    Set baRegion = viewerSheet.[I1].CurrentRegion
    
    If coRegion.Rows.Count < 3 Then Exit Sub
    If baRegion.Rows.Count < 3 Then Exit Sub

    
    '对表单的项目进行除外
    Call MakeExceptionRow

    '生成空行填充不平项目
    Call MakeRowsEven

    '检查完毕，开始对账
    '进行单项对单项核对
    Dim iRow As Integer
    Dim jRow As Integer
    Dim RemCo As Range
    Dim RemBa As Range
    Dim RowCo As Range
    Dim RowBa As Range
    '第一次循环，寻找一项配平项目
    
    Dim intTimesFound As Integer '统计找到几次
    Dim MatchAddress As String  ' 记录找到的匹配地址
    Dim SwitchAddres As String  ' 记录需要对调的匹配地址
    For iRow = 3 To coRegion.Rows.Count
        intTimesFound = 0 '找到几次计数器归零
        MatchAddress = ""
        SwitchAddres = ""
        
        Set RemCo = viewerSheet.Range("$F$" & iRow)
        For jRow = 3 To baRegion.Rows.Count
            Set RowCo = RemCo.Offset(0, -5).Resize(1, 7)
            Set RemBa = viewerSheet.Range("$N$" & jRow)
            Set RowBa = RemBa.Offset(0, -5).Resize(1, 7)
            If RemCo.Interior.Color = rgbWhite And _
                    RemBa.Interior.Color = rgbWhite Then
                    
                If RemCo.Value = RemBa.Value Then '如果贷方余额相等
                    '以公司为准对齐两行，并标记颜色
                    '记录需要调换位置的单元格位置
                    If MatchAddress = "" Then
                        MatchAddress = RemBa.Address
                    Else
                        MatchAddress = _
                        Join(Array(MatchAddress, RemBa.Address), ",")
                    End If
                    ' 被调换的为公司方与银行方对齐的位置 + intTimesFound 的行间offset
                    If SwitchAddres = "" Then
                        SwitchAddres = RemCo.Offset(intTimesFound, 8).Address
                    Else
                        SwitchAddres = _
                        Join(Array(SwitchAddres, _
                        RemCo.Offset(intTimesFound, 8).Address), ",")
                    End If

                    intTimesFound = intTimesFound + 1
                    'Call MakeRelation(RemCo, RemBa, intTimesFound)
                    'Call PointArrow(RemCo, RemBa)
                    '将中心格子的内容变为= 会使两表相连currentregion出错
                End If
            End If
        Next jRow
        
        ' 如果找到了至少一项，需要
        ' 1. 将RowCo标记
        ' 2. 将记录的行进行对调
        ' [Obsolete] 3. 在RowCo之后插入intTimesFound - 1 行
        ' 3. 在公司方 rowco开始之后寻找第一个rgbGray的格子
        ' 执行对调
        Dim firstGrayRange As Range
        i = coRegion.Rows.Count
        Do While i >= 3
            Set rng = viewerSheet.Range("$F$" & i)
            If rng.Interior.Color = rgbGray And rng.Offset(-1, 0).Interior.Color <> rgbGray Then
                Set firstGrayRange = rng
                Exit Do
            End If
            i = i - 1
        Loop

        If intTimesFound > 0 Then
            If intTimesFound = 1 Then
                
                Call MarkRowStatus(RowCo, EntryStatus.dzCertain)
            Else
                Call MarkRowStatus(RowCo, EntryStatus.dzPossibleMatch)
            End If
            Call SwitchRow(MatchAddress, SwitchAddres)
            Do While intTimesFound > 1
                'RowCo.Offset(1, 0).Insert (xlDown)
                ' 改为与 firstGrayRange的offset对调
                Call SwitchRow(RowCo.Offset(1, 0).Address, _
                    firstGrayRange.Offset(intTimesFound - 1, 0).Address)
                With RowCo.Offset(1, 0)
                    .Formula = emptyFormula
                    .Interior.Color = rgbLightGray
                    .Font.Color = rgbWhite
                End With
                intTimesFound = intTimesFound - 1
            Loop
        End If
    Next iRow
    '对右表未对齐项进行排列组合
    'Call CombineUnconsolidatedRows
End Sub
' 使target 与 toRow 位置的行调换位置
Function SwitchRow_Obsolete(target As Range, toRow As Integer)
    Dim temp As Variant
    Dim fromRow As Integer
    Dim destination, origin As Range
    '扩展target至整行
    fromRow = target.Row
    
    Set destination = target.CurrentRegion.Rows(toRow)
    Set origin = target.CurrentRegion.Rows(fromRow)
    
    temp = origin.Cells.Formula
    origin.Cells.Formula = destination.Cells.Formula
    destination.Cells.Formula = temp

End Function
' 使 Origin 与 Destin 位置的行调换位置, 两者代表了某格的 Address
Function SwitchRow(origin As String, Destin As String)
    Dim temp As Variant
    Dim rngOrig As Range
    Dim rngDest As Range
    Dim Orig As Range
    Dim Dest As Range
    Set rngOrig = viewerSheet.Range(origin)
    Set rngDest = viewerSheet.Range(Destin)
    
    ' 检查两者的数量是否相等
    If rngOrig.Count <> rngDest.Count Then Exit Function
    ' 这里不可以按照数列的序号取，如果中间有间隙的话它
    ' 代表了它的下一个紧贴着的单元格
    ' Destin恰好是连续的所以没有产生错误
    Dim RowsFound As Integer
    RowsFound = rngOrig.Count
    i = 1
    For Each rng In rngOrig
        Set Orig = rng.Offset(0, -5).Resize(1, 7)
        Set Dest = rngDest.Cells(i).Offset(0, -5).Resize(1, 7)
        temp = Orig.Cells.Formula
        Orig.Cells.Formula = Dest.Cells.Formula
        Dest.Cells.Formula = temp
        If RowsFound > 1 Then
            Call MarkRowStatus(Dest, EntryStatus.dzPossibleMatch)
        Else
            Call MarkRowStatus(Dest, EntryStatus.dzCertain)
        End If
        i = i + 1
    Next rng
    
End Function
' 将某个行标记为【已经匹配】
Function MarkRowStatus(target As Range, matchStatus As Long)
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

'将含有特定关键字的条目除外
Sub MakeExceptionRow()
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(年初恢复零余额账户用款额度.*|本月合计.*|本年累计.*|上年结转.*)"
    
    '构造C:C与K:K的有内容的单元格
    Dim rngZhaiyaoLeft, rngZhaiyaoRight, rng As Range
    Set rngZhaiyaoLeft = Intersect(viewerSheet.[C:C] _
            , viewerSheet.[C2].CurrentRegion)
    Set rngZhaiyaoRight = Intersect(viewerSheet.[K:K] _
            , viewerSheet.[K2].CurrentRegion)
            
    For Each rng In Union(rngZhaiyaoLeft, rngZhaiyaoRight)
        If regex.test(rng.Value) Then
            With rng.Offset(0, -2).Resize(1, 7)
                .Interior.Color = rgbGray
                .Font.Color = rgbYellow
            End With
            
        End If
    Next rng
End Sub
'带入某一方的任意单元格进行排列组合
Sub CombineUnconsolidatedRows(target As Range)

End Sub

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

' 该子程序会寻找CoRegion和BaRegion中单侧有颜色的项目，并且用空行填充直至齐平
Sub MakeRowsEven()
    '寻找两个区域中行数多的一个作为max
    Dim MaxRow As Integer
    Dim coRow As Range
    Dim baRow As Range
    MaxRow = WorksheetFunction.Max(coRegion.Rows.Count, baRegion.Rows.Count)

    ' 循环MaxRow到最后一行，单侧颜色有，另一侧无时，将无的那一侧设为空行并打上灰色
    ' 定义空行的formula

    Set coRow = coRegion.Rows(3)
    Set baRow = baRegion.Rows(3)
    Dim MatchColor As Long
    ' 将匹配颜色初始化成rgbGray
    MatchColor = rgbGray
    Do While coRow.Cells(1).Value <> "" Or baRow.Cells(1).Value <> ""
        ' 如果颜色匹配，则生成新的匹配颜色
        If coRow(1).Interior.Color = _
            baRow(1).Interior.Color Then
            MatchColor = coRow(1).Interior.Color

        End If
        ' 如果当前颜色不等于匹配颜色，则在原地增加空白行
        If coRow(1).Interior.Color <> MatchColor Then
            coRow.Insert (xlDown)
            With coRow.Offset(-1, 0)
                .Formula = emptyFormula
                .Interior.Color = rgbLightGray
                .Font.Color = rgbWhite
            End With
            '增加空白行后需往上缩进
            Set coRow = coRow.Offset(-1, 0)
        End If

        If baRow(1).Interior.Color <> MatchColor Then
            baRow.Insert (xlDown)
            With baRow.Offset(-1, 0)
                .Formula = emptyFormula
                .Interior.Color = rgbLightGray
                .Font.Color = rgbWhite
            End With
            '增加空白行后需往上缩进
            Set baRow = baRow.Offset(-1, 0)
        End If

        Set coRow = coRow.Offset(1, 0)
        Set baRow = baRow.Offset(1, 0)
    Loop

End Sub








