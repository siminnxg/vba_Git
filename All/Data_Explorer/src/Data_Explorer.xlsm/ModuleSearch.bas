Attribute VB_Name = "ModuleSearch"
Option Explicit

Public Sub SaveSearch()
    
    Dim varSearchCol As Variant '---�˻��� �� ����
    Dim strSearchData As String '---�Էµ� �˻��� ����
    Dim rngSearch As Range
    Dim rngKeyword As Range
    Dim varPresetIdx As Variant
    
    Call SetRange
    
    If ����������.Value = Empty Then
        Exit Sub
    End If
    
    Set rngKeyword = Range(�˻�Ű����_����, �˻�Ű����_��)
    Set rngSearch = rngKeyword.Offset(-1, 0)
    
    varSearchCol = rngSearch.Cells.Count
    
    '�˻��Ǿ� �ִ� ���� ���� �� ������ ����
    strSearchData = rngKeyword(1) & "��" & rngSearch(1) '---�����ڷ� �� ���
    
    For i = 2 To varSearchCol
        
        strSearchData = strSearchData & "��" & rngKeyword(i) & "��" & rngSearch(i)
        
    Next
    
    If IsNull(strSearchData) = False Then
        With Sheets("etc")
            varPresetIdx = .Range("preset_list").Find(what:=����������.Value, lookat:=xlWhole).Row '---etc ��Ʈ������ ������ ��ġ ã��
    
            '�˻����̴� ������ �����¿� �����ϱ�
            .Cells(varPresetIdx, 7) = strSearchData
    
        End With
    End If
    
End Sub

Public Sub LoadSearch()
    
    Dim varSearchData As Variant
    Dim varPresetIdx As Variant
    Dim rngSel As Range
    
    On Error Resume Next
    
    Call SetRange
    
    With Sheets("etc")
    
        varPresetIdx = .Range("preset_list").Find(what:=����������.Value, lookat:=xlWhole).Row '---etc ��Ʈ������ ������ ��ġ ã��
        varSearchData = Split(.Cells(varPresetIdx, 7), "��")
    
    End With
    
    '����� �� ������ ����
    If UBound(varSearchData) = -1 Then
        Exit Sub
    End If
    
    Call UpdateStart
    Call SetColor
    
    For i = 0 To UBound(varSearchData)
        j = 0
        
        If varSearchData(i + 1) <> "" Then
            For Each rngSel In �˻�Ű����
            
                If rngSel.Value = varSearchData(i) Then
                    
                    rngSel.Offset(-1, 0).Value = varSearchData(i + 1)
                    rngSel.Interior.Color = colorCategorySel
                    
                    i = i + 1
                    Exit For
                    
                End If
                
                j = j + 1
            Next
        End If
    Next
    
    Call UpdateEnd
End Sub

Public Sub PasteData()
    
    Dim rngData As Range
    Dim varDataAry As Variant
    
    Call UpdateStart
    Call SetRange
    
    Set rngData = Sheets(����������.Value).ListObjects(1).Range
    
    Application.CutCopyMode = True
    
    With rngData.SpecialCells(xlCellTypeVisible)

        .Copy
        �˻�Ű����_����.PasteSpecial xlPasteValues

    End With

    Application.CutCopyMode = False
    
    'DATA �̸� ���� �缳��
    ThisWorkbook.Names("DATA").RefersTo = Range(Selection, �˻���_����)
    
    'DATA �̸� ������ ���Ǻ� ���� ����
    With Range("DATA").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$F$5<>""""")
        
        '�� �׵θ�
        .Borders(xlLeft).LineStyle = xlContinuous
        .Borders(xlRight).LineStyle = xlContinuous
        .Borders(xlTop).LineStyle = xlContinuous
        .Borders(xlBottom).LineStyle = xlContinuous
        
    End With
    
    Sheets("Search").Range("A1").Select
    
    Call UpdateEnd
End Sub


'=====================================================================
'��ũ�� : ResetSearch
'��� ��Ʈ : Home ��Ʈ
'���� : �˻� ���� �˻� �ʱ�ȭ ��ư, �˻����̴� ���� �ʱ�ȭ �� ������ ��Ʈ ���� ����
'=====================================================================
Public Function ResetSearch()
    
    'ResetSearch = 1 : ���õ� ���� ����
    'ResetSearch = 0 : ���� ����
        
    '---�������� ����ϴ� ���� ��ġ ȣ��
    Call SetRange
    
    '---�˻��� �ʱ�ȭ
    Range(�˻���_����, �˻�Ű����_��.Offset(-1, 0)).ClearContents
    
    '---���õ� �� �� ���� �ʱ�ȭ
    Range(�˻�Ű����_����, �˻�Ű����_��).ClearFormats
    
    '---���õ� ���� �������� ������ ���� ����
    If �˻�Ű����_���� = "" Then
    
        Range("notice") = "���õ� ī�װ��� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        ResetSearch = 1
        
        Exit Function
        
    End If
    
    '---������ ��Ʈ ������ ���Ͱ� �ɷ��ִٸ� ����
    If Sheets(CStr(����������.Value)).AutoFilter.FilterMode = True Then

        Sheets(����������.Value).ShowAllData

    End If
    
    ResetSearch = 0
    
End Function

Public Sub AutoFill(varIndex As Variant)
    
    Dim rngStart As Range
    Dim objData As Object
    Dim varSelCol As Variant
    
    Set objData = Sheets(����������.Value).ListObjects(1)
    varSelCol = objData.ListColumns(varIndex).Index
    
    Set rngStart = Sheets(����������.Value).Columns(varSelCol).Cells(1)
        
    Do Until rngStart.Row >= objData.Range.Rows.Count
        
        If rngStart.Offset(1, 0) = "" Then
            Range(rngStart, rngStart.End(xlDown).Offset(-1, 0)).FillDown
        
        End If
        
        Set rngStart = rngStart.End(xlDown)
        
    Loop
    
    Call PasteData

End Sub

