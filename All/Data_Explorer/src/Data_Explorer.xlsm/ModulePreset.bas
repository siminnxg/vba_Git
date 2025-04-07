Attribute VB_Name = "ModulePreset"
'=====================================================================
'��ũ�� : SavePreset
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ����ڰ� �Է��� ���� ������ ���������� ����
'=====================================================================
Public Sub SavePreset()
    
    '###���� ����###
    
    Dim preset_count As Variant
    Dim file_info As Variant
        
    '###���� ����###
        
    '---file_info �迭�� ���� ���� ����
    file_info = Array(preset, File_adr, File_name, sheet_name)
    
    '---��ĭ ���� �� �˸�â ǥ��
    If File_adr = "" Or File_name = "" Or sheet_name = "" Or preset = "" Then
    
        MsgBox "���� ��� �Է����ּ���."
        Exit Sub
        
    End If
        
    With Sheets("etc")
        '---������ �����¸� ���� �� �˸�â ǥ��
        If Not .Range("B:B").Find(what:=preset, lookat:=xlWhole) Is Nothing Then
        
            MsgBox "������ ������ ���� �����մϴ�."
            Exit Sub
            
        End If
        
        '---������ ���� Ȯ��
        If .Range("B2") = "�������" Then
        
            preset_count = 1
            
        Else
        
            preset_count = .Range("B1").End(xlDown).Row
            
        End If
                    
        '---�����¿� ���� ���� ����
        For i = 0 To 3
        
            .Cells(preset_count + 1, 2 + i).Value = file_info(i)
            
        Next
        
        '---preset_list �̸� ���� ������
        ThisWorkbook.Names("preset_list").RefersTo = .Range("B1", .Cells(preset_count + 1, 2))
    
    End With
    
    '---�����̼� ���ΰ�ħ
    ActiveWorkbook.SlicerCaches("�����̼�_������").PivotTables(1).PivotCache.Refresh
    
End Sub
 
'=====================================================================
'��ũ�� : LoadPreset
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ���õ� ������ ������ �����ͼ� ������ ȣ��
'=====================================================================
Public Sub LoadPreset()
    
    '###���� ����###
    
    Dim varPresetSel As Variant '---���� ���õ� ������ �� ���� ����
    Dim varSearchRow As Variant '---etc ��Ʈ���� ���� ���õ� �������� ��ġ ���� ����
    Dim varCategoryAdr As Variant '---�����¿� ����Ǿ��ִ� ���õ� ������ ��ġ ���� ����
        
    '###���� ����###
    
    'ȭ�� ������Ʈ ����
    Call UpdateStart
    
    '�������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call SetRange
    Call SetColor
    
    '������ �߻��ϸ� ����
    
    On Error GoTo exit_error
    
    With Sheets("etc").PivotTables("������").DataBodyRange
        
        varPresetSel = CStr(.Cells(1)) '---���� ���õ� ������ �� ������ ����
        
        '������ ���� ����ó��
        If varPresetSel = "Preset_Header" Then
            
            MsgBox ("�������� �������ּ���,")
            GoTo exit_sub
                            
        ElseIf varPresetSel = "�������" Then
            
            MsgBox ("�������� �������� �ʽ��ϴ�.")
            GoTo exit_sub
        
        ElseIf .Cells.Count > 1 Then
            
            MsgBox ("�������� 1���� �������ּ���.")
            GoTo exit_sub
            
        End If
    End With
    
    Call HideSearchSht(False)
    
    '���� �˻� ������ ���� �� �ʱ�ȭ
    Call SaveSearch
    Call ClearHomeData
    
    With Sheets("etc")
        'etc ��Ʈ������ ������ ��ġ ã��
        varSearchRow = .Range("preset_list").Find(what:=varPresetSel, lookat:=xlWhole).Row
        
        '������ �� �ٿ��ֱ�
        �����¸� = .Cells(varSearchRow, 2).Value
        ���ϰ�� = .Cells(varSearchRow, 3).Value
        ���ϸ� = .Cells(varSearchRow, 4).Value
        ��Ʈ�� = .Cells(varSearchRow, 5).Value
        varCategoryAdr = .Cells(varSearchRow, 6).Value
        
    End With
    
    '��Ʈ ��� ���� ��Ӵٿ� ����
    ��Ʈ��.Validation.Delete
    
    Call SearchCategory(varPresetSel)
    
    '�� ���� ���¸� �����س��� ��� ȣ��
    If Not varCategoryAdr = "" Then
    
        Sheets("Search").Range(varCategoryAdr).Interior.Color = colorCategorySel
        
    End If
    
    '���õ� �� ����
    Call AddCategory
    
'---���� ó��
exit_sub:
    
    Call UpdateEnd
    Exit Sub
    
'---���� �߻� ó��
exit_error:
    
    MsgBox ("������ �߻��߽��ϴ�. ���� �ڵ� : " & Err.Number & " " & Err.Description)
    Call UpdateEnd
End Sub

'=====================================================================
'��ũ�� : DeletePreset
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ���õ� �������� ����
'=====================================================================
Public Sub DeletePreset()

    Dim varPresetSel As String
    Dim varSearchRow As Variant
    
    Call UpdateStart
    Call SetRange
    
    '---���� �߻��ص� �����ϰ� ����
    On Error Resume Next
    
    '---��Ʈ ���� �� �ý��� ���� �̳���
    Application.DisplayAlerts = False
    
    With Sheets("etc").PivotTables("������").DataBodyRange
        
        '---���õ� �������� ���� �� ó��
        If .Cells(1) = "Preset_Header" Then
        
            MsgBox ("�������� �������ּ���.")
            GoTo exit_sub
        
        '---�������� �������� ���� �� ó��
        ElseIf .Cells(1) = "�������" Then
        
            MsgBox ("�������� �������� �ʽ��ϴ�.")
            GoTo exit_sub
        
        '---2�� �̻� ������ ���� �� ó��
        ElseIf .Cells.Count > 1 Then
        
            If MsgBox("�������� �ΰ� �̻� ���õǾ����ϴ�." & vbCrLf & "��� �����Ͻðڽ��ϱ�?", vbYesNo) = vbNo Then
            
                GoTo exit_sub
                
            End If
        End If
        
        '---���õ� ������ ������ŭ �ݺ�
        For i = 1 To .Cells.Count
            
            '--���� ������ ��Ͽ� ���� ��ȸ ���� �����Ͱ� �ִ� ��� ��ȸ���� ������ �ʱ�ȭ
            If CStr(.Cells(i)) = ����������.Value Then
                
                '---�˻� ���� �ʱ�ȭ
                Call ClearHomeData
                
                '---�� ���� ���� �ʱ�ȭ
                �����.Clear
                ����������.ClearContents
                
                Call HideSearchSht(True) '---�˻� ���� �����
            End If
            
            '---������ ���� �� ��Ʈ, ���� �Բ� ����
            Sheets(CStr(.Cells(i))).Visible = True
            Sheets(CStr(.Cells(i))).Delete
            ActiveWorkbook.Queries(CStr(.Cells(i))).Delete
            
            '---�����Ϸ��� �����¸� etc ��Ʈ���� ��ġ �˻�
            varSearchRow = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc ������ ����Ʈ �������� ������ ���� ����
            Range(Range("preset_list")(varSearchRow), Range("preset_list")(varSearchRow).Offset(0, 5)).Delete Shift:=xlUp
            
            '---�����ִ� ������ ���� ��� ó��
            If Range("preset_list").Cells.Count = 1 Then
                
                Range("Preset_list").Offset(1, 0) = "�������"
                ThisWorkbook.Names("preset_list").RefersTo = Sheets("etc").Range("B1:B2") '---preset_list �̸� ���� ������
                
                '---���� ��ü ����
                Call DeleteConnect
                
            End If
        Next
    End With
    
    '---�����̼� ������ ���ΰ�ħ
    ActiveWorkbook.SlicerCaches("�����̼�_������").PivotTables(1).PivotCache.Refresh
    
    '---������ �̸� ���� ������
    ThisWorkbook.Names("DATA").RefersTo = �˻���_����
    
    '---�ý��� ���� ����
    Application.DisplayAlerts = True

'���� ó��
exit_sub:
    Call UpdateEnd
    
End Sub


Public Sub EditPreset()
    
    Dim strFileAdr As String
    Dim varSearchRow As Variant
        
    Call UpdateStart
    
    With Sheets("etc").PivotTables("������").DataBodyRange
        
        '---���õ� �������� ���� �� ó��
        If .Cells(1) = "Preset_Header" Then
        
            MsgBox ("�������� �������ּ���.")
            GoTo exit_sub
        
        '---�������� �������� ���� �� ó��
        ElseIf .Cells(1) = "�������" Then
        
            MsgBox ("�������� �������� �ʽ��ϴ�.")
            GoTo exit_sub
            
        End If
        
        '---��� �Է� �ڽ� ǥ��
        strFileAdr = InputBox("������ ���� ��θ� �Է����ּ���.", "������ ��� ����")
        
        '---�Է��� ���� ���� ��� ����
        If strFileAdr = Empty Then
            
            GoTo exit_sub
            
        End If
        '---���õ� ������ ������ŭ �ݺ�
        For i = 1 To .Cells.Count
            
            '---�����Ϸ��� �����¸� etc ��Ʈ���� ��ġ �˻�
            varSearchRow = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc ������ ����Ʈ �������� ������ ��� ����
            Range("preset_list")(varSearchRow).Offset(0, 1).Value = strFileAdr
            
        Next
        
    End With
    
'���� ó��
exit_sub:
    Call UpdateEnd
    
End Sub


Public Sub RefreshPreset()
    
    Call SetRange
    
    With Range("preset_list")
        If .Cells.Count < 2 Then
            
            MsgBox "�������� �������� �ʽ��ϴ�."
            Exit Sub
            
        End If
        
        For i = 2 To .Cells.Count
            '�Է��� ��ο� ���� ���� üũ
            If CheckFile(���ϰ�� & "\" & ���ϸ�) = False Then
                
                MsgBox (���ϰ�� & " ��ο� " & ���ϸ� & " ������ �������� �ʽ��ϴ�.")
            
            Else
                
                '������ �̸����� ������ ��Ʈ�� listobject ���ΰ�ħ
                Sheets(CStr(.Cells(i).Value)).ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
            
            End If
        Next
        
        MsgBox "�ֽ� �����͸� �����Ͽ����ϴ�."
        
    End With
End Sub
