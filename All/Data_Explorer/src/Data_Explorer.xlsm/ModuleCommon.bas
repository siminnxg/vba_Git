Attribute VB_Name = "ModuleCommon"
'=====================================================================
'��Ÿ ��� ���
'=====================================================================

'###���� ����###

Public File_name As String
Public File_adr As String
Public preset As String
Public sheet_name As String

'###���� ���� ����###

Public user_file_adr As Range
Public user_file_name As Range
Public user_file_sheet As Range
Public user_file_preset As Range

Public act_sheet_name As Range
Public act_category_list As Range
Public act_category_start As Range
Public act_category_end As Range

Public search_user_start As Range
Public search_category_start As Range
Public search_FixRow As Range
Public search_category_end As Range

Public etc_preset As Range

'=====================================================================
'ȭ�� ������Ʈ ���� (���� �ӵ� ����)
Public Sub update_start()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub

'=====================================================================
'ȭ�� ������Ʈ ����
Public Sub update_end()
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

'=====================================================================
'������ ��Ʈ�� üũ
Public Function query_check()
    
    Dim sheet_count As Variant '---��Ʈ ���� ���� ����
        
    sheet_count = ActiveWorkbook.Sheets.Count '---���� ���Ͽ��� ��Ʈ ���� üũ
    
    '---��Ʈ ������ŭ �ݺ�
    For i = 1 To sheet_count
        
        If preset = ActiveWorkbook.Sheets(i).Name Then '---�Է��� ������ ���� ���� �����Ǿ��ִ� ��Ʈ�� ������ �̸����� Ȯ��
        
            query_check = 1
        
        End If
    Next
    
End Function

'=====================================================================
'�� ����Ʈ ȣ�� ���� Ȯ��
Public Function category_check()
    
    '---�� ����Ʈ ���� ù��° �� ���� üũ
    If act_category_start.Value = "" Then
    
        category_check = 1
        
    End If
    
End Function

'=====================================================================
'�� ��� ���� ����
Public Sub range_set()
    
    With Sheets("home")
                
        Set user_file_adr = .Range("C4") '---���� ���
        Set user_file_name = .Range("C5") '---���� �̸�
        Set user_file_sheet = .Range("C6") '---��Ʈ ���
        Set user_file_preset = .Range("C7") '---������ �̸�
        
        Set act_sheet_name = .Range("G4") '---���� ������ �̸�
        Set act_category_start = act_sheet_name.Offset(1, 0) '---�� ����Ʈ ó�� ��ġ
        Set act_category_end = act_sheet_name.End(xlDown) '---�� ����Ʈ ������ ��ġ
        Set act_category_list = Range(act_category_start, act_category_end) '---�� ����Ʈ ��ü ����
        
        Set search_user_start = .Range("K4") '---�˻� ���� ����
        Set search_category_start = .Range("K5") '---���õ� �� ���� ����
        Set search_FixRow = .Range("J8") '---�� ���� �Է� ����
        
        '---���õ� �� �� ����
        If search_category_start = "" Then
            
            Set search_category_end = search_category_start
            
        Else
        
            Set search_category_end = search_category_start.Offset(0, -1).End(xlToRight)
            
        End If
        
    End With
    
    With Sheets("etc")
    
        Set etc_preset = .Range("H2") '---������ ���� ����
    
    End With
    
    '---sub ���� ������ �� ����
    Call file_info_load

End Sub

'=====================================================================
'����ڰ� �Է��� ���� ���� ������ ����
Public Sub file_info_load()
    
    File_adr = user_file_adr.Value '---���� ��� ����
    File_name = user_file_name.Value '---���� �̸� ����
    sheet_name = user_file_sheet.Value '----��Ʈ �̸� ����
    preset = user_file_preset.Value '---������ �̸� ����
    
End Sub

'=====================================================================
'���� �˻����� ����, ī�װ� �ʱ�ȭ
Public Sub home_data_clear()
    
    '---���� �ҷ��� �����Ͱ� ������ ����
    If act_sheet_name.Value = Empty Then
        
        Exit Sub
    
    End If
    
    '--- sub �˻����� ���� �ʱ�ȭ
    search_reset
    
    Range(search_user_start, search_category_end).Clear
    
    Range("DATA").ClearContents '---�� ���� ���� �ʱ�ȭ
    
    Range("DATA").FormatConditions.Delete
    
    '--- sub �� ����Ʈ ���� �ʱ�ȭ
    Call category_reset
    
    Range("notice").ClearContents
    
    
End Sub
      
'=====================================================================
'���� �ѹ��� ����
Public Sub connect_delete()

    Dim conn As Object
    Dim connName As String
    
    '---��� ������ ��ȸ
    For Each conn In ActiveWorkbook.Connections
        connName = conn.Name '---���� �̸� ��������
        
        '---����� �����ϴ� �̸��� ���� ����
        If connName Like "����*" Then
        
            conn.Delete
            
        End If
    Next conn
    
End Sub

'=====================================================================
'���� ���� ���� üũ
Public Function FileExists(ByVal path_ As String) As Boolean
        
    FileExists = (Dir(path_, vbDirectory) <> "") '---�Էµ� ��ο� �Էµ� ���ϸ��� �����ϴ��� Ȯ��
 
End Function

'=====================================================================
'���� ���� ���� üũ
Public Function IsCheckOpen(CheckFile As String) As Boolean

    Dim wb As Variant
    
    On Error Resume Next
    
    Set wb = Workbooks(CheckFile)
        
    If Not wb Is Nothing Then
    
        IsCheckOpen = True
        
    Else
    
        IsCheckOpen = False
        
    End If
    
    On Error GoTo 0
    
End Function

'=====================================================================
'������ �̸� üũ
Public Function preset_name_check()
    
    Dim preset_name_index As Variant '---�����¸� ����
    Dim check As Boolean '---������ �����¸� üũ
    
    preset_name_index = 1
    
    With Range("preset_list")
        Do
            For i = 2 To .Cells.Count
                
                '---������ ������ �̸��� �����ϴ� ��� ó��
                If StrComp(.Cells(i).Value, "������" & preset_name_index) = 0 Then
                
                    preset_name_index = preset_name_index + 1
                    check = True
                    Exit For
                    
                End If
                
                check = False
            Next
            
            '---������ ������ �̸��� ������ ����
            If check = False Then
            
                Exit Do
                
            End If
        Loop
    End With
    
    '---������ �̸� ��ȯ
    preset_name_check = "������" & preset_name_index
    
End Function
