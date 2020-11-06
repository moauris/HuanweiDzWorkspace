Option Explicit
' HuanweiDZ Macro �汾 1.0.0.2 alpha
' ģ�飺������
' �汾���ڣ�2020-11-07
' ���ߣ�https://github.com/moauris
' ��ϵ��ʽ��mchenf@icloud.com
 
 ' һЩ�����������
Dim i As Integer
Dim j As Integer
Dim rng As Range

Dim viewerSheet As Worksheet '�����˱����ʾ��
Private coRegion As Range, baRegion As Range '�����˹�˾�����з���������
Dim emptyFormula As Variant 'ȫ�ֱ��������ڱ�ʾһ�п�ֵ

' ���ڸ�ʽ�ض��е�״̬��ʾ��ö��
Public Enum EntryStatus
    dzUnmatched = 0
    dzException = 1
    dzPossible = 2
    dzCertain = 3
    dzFiller = 4
End Enum

' ���������
Sub Main()
    ' ����ȫ�ֱ��� viewerSheet
    Set viewerSheet = ThisWorkbook.Worksheets("�����ʾ��")
    ' ����ȫ�ֱ��� emptyFormula
    emptyFormula = Array("'_", "'_", "'_", 0, 0, 0, 0)
    
    ' �����ļ�
    Call Step010_ImportSourceFiles

    ' �ж� viewerSheet ״̬�Ƿ���������״̬
    If Step011_IsViewerIncomplete Then Exit Sub

    ' ���е��Ե��Ķ���
    Call Step020_ConsolidateSingle
    
    ' ���ջ���񣬰���һЩ��ʽ������
    Call Step999_Finalize

End Sub

