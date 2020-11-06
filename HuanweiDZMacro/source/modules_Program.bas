Option Explicit
' HuanweiDZ Macro 版本 1.0.0.2 alpha
' 模块：主程序
' 版本日期：2020-11-07
' 作者：https://github.com/moauris
' 联系方式：mchenf@icloud.com
 
 ' 一些常规变量声明
Dim i As Integer
Dim j As Integer
Dim rng As Range

Dim viewerSheet As Worksheet '代表了表格显示区
Private coRegion As Range, baRegion As Range '代表了公司、银行方工作区域
Dim emptyFormula As Variant '全局变量，用于表示一行空值

' 用于格式特定行的状态表示的枚举
Public Enum EntryStatus
    dzUnmatched = 0
    dzException = 1
    dzPossible = 2
    dzCertain = 3
    dzFiller = 4
End Enum

' 主程序入口
Sub Main()
    ' 定义全局变量 viewerSheet
    Set viewerSheet = ThisWorkbook.Worksheets("表格显示区")
    ' 定义全局变量 emptyFormula
    emptyFormula = Array("'_", "'_", "'_", 0, 0, 0, 0)
    
    ' 导入文件
    Call Step010_ImportSourceFiles

    ' 判定 viewerSheet 状态是否满足运行状态
    If Step011_IsViewerIncomplete Then Exit Sub

    ' 进行单对单的对账
    Call Step020_ConsolidateSingle
    
    ' 最终化表格，包括一些格式和清理
    Call Step999_Finalize

End Sub

