Option Explicit
' HuanweiDZ Macro 版本 1.0.0.3 alpha
' 版本日期：2020-11-09
' 作者：https://github.com/moauris
' 联系方式：mchenf@icloud.com
' 版本更新内容 1.0.0.3
' viewerSheet 中去掉了贷方余额一项
' 在对账开始之前需要增加一条有限调整项，对余额低的一方进行余额的对齐
' 将所有不引用 viewerSheet 的通用函数放在了这里
 
Dim viewerSheet As Worksheet '代表了表格显示区
Dim coRegion, baRegion As Range '代表了公司、银行方工作区域
Dim arrowColor As Long '全局变量箭头颜色
Dim i As Integer
Dim j As Integer
Dim rng As Range

Dim emptyFormula(1 To 6) As Variant '全局变量，用于表示一行空值
' Enumeration for the Status of the Entries
Public Enum EntryStatus
    dzUnmatched = 0
    dzException = 1
    dzPossibleMatch = 2
    dzCertain = 3
End Enum
Public Enum Sides
    Company = 0
    Bank = 1
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
    viewerSheet.[G:H].Clear
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
    ' 初始化空的方程列（6）
    emptyFormula(1) = "'-"
    emptyFormula(2) = "'-"
    emptyFormula(3) = "'-"
    emptyFormula(4) = 0
    emptyFormula(5) = 0
    emptyFormula(6) = 0
    If targetSheet Is Nothing Then Exit Sub
    LedgerTitle = targetSheet.[A1].Value
    
    'Debug.Print (LedgerTitle)
    '辅助明细账  = Bank
    '科目明细账 = Company
    Select Case Left(LedgerTitle, 5)
        Case "辅助明细账"
            Call SyncToViewer(targetSheet, viewerSheet, Sides.Bank)
        Case "科目明细账"
            Call SyncToViewer(targetSheet, viewerSheet, Sides.Company)
        Case Else
            '没有找到合适的抬头文字，直接退出
            Dim message As String
            message = "请检查导入的表格是否是一个合适的文件，或者抬头是否为“辅助明细账”或者“科目明细账”，并且不含有特殊符号"
            MsgBox message, vbCritical, "表格不符合导入规范"
            GoTo CLEAN_UP
    End Select

    Call TryConsolidateSingle
    '导入和单项对账完成，标记双方工作区
    Dim CurrencyColumns As Range
    '对账完成，开始制造中文货币格式
    Set CurrencyColumns = viewerSheet.[D:F,L:N]
    CurrencyColumns.NumberFormat = _
        "_ [$￥-zh-CN]* #,##0.00_ ;_ [$￥-zh-CN]* -#,##0.00_ ;_ [$￥-zh-CN]* ""-""??_ ;_ @_ "
CLEAN_UP:
    Set targetSheet = Nothing
End Sub

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
        MatchAddress = "" '匹配地址归零
        SwitchAddres = "" '交换地址归零
        
        Set RemCo = viewerSheet.Range("$F$" & iRow)
        For jRow = 3 To baRegion.Rows.Count
            Set RowCo = RemCo.Offset(0, -5).Resize(1, 7)
            Set RemBa = viewerSheet.Range("$N$" & jRow)
            Set RowBa = RemBa.Offset(0, -5).Resize(1, 7)
            If RemCo.Interior.Color = rgbWhite And _
                    RemBa.Interior.Color = rgbWhite Then
                '如果贷方余额相等
                '以公司为准对齐两行，并标记颜色
                '记录需要调换位置的单元格位置
                If RemCo.Value = RemBa.Value Then
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
        For i = coRegion.Rows.Count To 3 Step -1
            Set rng = viewerSheet.Range("$F$" & i)
            If rng.Interior.Color = rgbGray And rng.Offset(-1, 0).Interior.Color <> rgbGray Then
                Set firstGrayRange = rng
                Exit For
            End If
        Next i

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

' 使 Origin 与 Destin 位置的行调换位置, 两者代表了某格的 Address
Sub SwitchRow(origin As String, Destin As String)
    Dim temp As Variant
    Dim rngOrig As Range
    Dim rngDest As Range
    Dim Orig As Range
    Dim Dest As Range
    Set rngOrig = viewerSheet.Range(origin)
    Set rngDest = viewerSheet.Range(Destin)
    
    ' 检查两者的数量是否相等
    If rngOrig.Count <> rngDest.Count Then Exit Sub
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