Attribute VB_Name = "ModulePreset"
'=====================================================================
'��ũ�� : preset_save
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ����ڰ� �Է��� ���� ������ ���������� ����
'=====================================================================
Public Sub preset_save()
    
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
'��ũ�� : preset_load
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ���õ� ������ ������ �����ͼ� ������ ȣ��
'=====================================================================
Public Sub preset_load()
    
    '###���� ����###
    
    Dim preset_select As Variant '---���� ���õ� ������ �� ���� ����
    Dim search_row As Variant '---etc ��Ʈ���� ���� ���õ� �������� ��ġ ���� ����
    Dim category_adr As Variant '---�����¿� ����Ǿ��ִ� ���õ� ������ ��ġ ���� ����
        
    '###���� ����###
    
    '---ȭ�� ������Ʈ ����
    Call update_start
    
    '---�������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call range_set
    Call color_set
    
    '---������ �߻��ϸ� ����
    On Error GoTo exit_error
    
    With Sheets("etc").PivotTables("������").DataBodyRange
    '---���� ���õ� ������ �� ������ ����
        preset_select = CStr(.Cells(1))
        
        If preset_select = "Preset_Header" Then
            
            MsgBox ("�������� �������ּ���,")
            GoTo exit_sub
                            
        ElseIf preset_select = "�������" Then
            
            MsgBox ("�������� �������� �ʽ��ϴ�.")
            GoTo exit_sub
        
        ElseIf .Cells.Count > 1 Then
            
            MsgBox ("�������� 1���� �������ּ���.")
            GoTo exit_sub
            
        End If
    End With
    
    '--- sub : �˻�, �� ���� ���� �ʱ�ȭ
    Call home_data_clear
    
    With Sheets("etc")
        '---etc ��Ʈ������ ������ ��ġ ã��
        search_row = .Range("preset_list").Find(what:=preset_select, lookat:=xlWhole).Row
        
        '---������ �� �ٿ��ֱ�
        user_file_preset = .Cells(search_row, 2).Value
        user_file_adr = .Cells(search_row, 3).Value
        user_file_name = .Cells(search_row, 4).Value
        user_file_sheet = .Cells(search_row, 5).Value
        category_adr = .Cells(search_row, 6).Value
        
    End With
    
    '---��Ʈ ��� ���� ��Ӵٿ� ����
    user_file_sheet.Validation.Delete
    
     '---�Է��� ��ο� ���� ���� üũ
    If FileExists(user_file_adr & "\" & user_file_name) = False Then
        
        MsgBox (user_file_adr & " ��ο� " & user_file_name & " ������ �������� �ʾ� ���� �������� �ҷ��ɴϴ�.")
    
    Else
        
        '---������ �̸����� ������ ��Ʈ�� listobject ���ΰ�ħ
        Sheets(preset_select).ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    
    End If
    
    Call search_list(preset_select)
    
    '---�� ���� ���¸� �����س��� ��� ȣ��
    If Not category_adr = "" Then
    
        Sheets("Home").Range(category_adr).Interior.Color = category_sel_color
        
    End If
    
    '---���õ� �� ����
    Call button_category_add
    
'---���� ó��
exit_sub:
    
    Call update_end
    Exit Sub
    
'---���� �߻� ó��
exit_error:
    
    MsgBox ("������ �߻��߽��ϴ�. ���� �ڵ� : " & Err.Number & " " & Err.Description)
    Call update_end
End Sub

'=====================================================================
'��ũ�� : preset_delete
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ���õ� �������� ����
'=====================================================================
Public Sub preset_delete()

    Dim preset_select As String
    Dim search_row As Variant
    
    Call update_start
    Call range_set
    
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
            If CStr(.Cells(i)) = act_sheet_name.Value Then
                
                '---�˻� ���� �ʱ�ȭ
                Call home_data_clear
                
                '---�� ���� ���� �ʱ�ȭ
                act_category_list.Clear
                act_sheet_name.ClearContents
                
                '---�� ���� ���� �����
                Call HideCategoryRng
                
            End If
            
            '---������ ���� �� ��Ʈ, ���� �Բ� ����
            Sheets(CStr(.Cells(i))).Delete
            ActiveWorkbook.Queries(CStr(.Cells(i))).Delete
            
            '---�����Ϸ��� �����¸� etc ��Ʈ���� ��ġ �˻�
            search_row = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc ������ ����Ʈ �������� ������ ���� ����
            Range(Range("preset_list")(search_row), Range("preset_list")(search_row).Offset(0, 5)).Delete Shift:=xlUp
            
            '---�����ִ� ������ ���� ��� ó��
            If Range("preset_list").Cells.Count = 1 Then
                
                Range("Preset_list").Offset(1, 0) = "�������"
                ThisWorkbook.Names("preset_list").RefersTo = Sheets("etc").Range("B1:B2") '---preset_list �̸� ���� ������
                
                '---���� ��ü ����
                Call connect_delete
                
            End If
        Next
    End With
    
    '---�����̼� ������ ���ΰ�ħ
    ActiveWorkbook.SlicerCaches("�����̼�_������").PivotTables(1).PivotCache.Refresh
    
    '---������ �̸� ���� ������
    ThisWorkbook.Names("DATA").RefersTo = search_user_start
    
    '---�˻� ���� �����
    Call home_data_hide
    
    '---�ý��� ���� ����
    Application.DisplayAlerts = True

'���� ó��
exit_sub:
    Call update_end
    
End Sub


Public Sub preset_edit()
    
    Dim strFileAdr As String
    Dim search_row As Variant
        
    Call update_start
    
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
            search_row = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc ������ ����Ʈ �������� ������ ��� ����
            Range("preset_list")(search_row).Offset(0, 1).Value = strFileAdr
            
        Next
        
    End With
    
'���� ó��
exit_sub:
    Call update_end
    
End Sub

