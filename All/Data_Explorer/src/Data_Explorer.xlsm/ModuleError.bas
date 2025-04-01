Attribute VB_Name = "ModuleError"
Option Explicit
'=====================================================================
'����ó�� ���� ���
'=====================================================================

'������ ��Ʈ�� üũ
Public Function CheckQuery()
    
    Dim sheet_count As Variant '---��Ʈ ���� ���� ����
        
    sheet_count = ActiveWorkbook.Sheets.Count '---���� ���Ͽ��� ��Ʈ ���� üũ
    
    '---��Ʈ ������ŭ �ݺ�
    For i = 1 To sheet_count
        
        If preset = ActiveWorkbook.Sheets(i).Name Then '---�Է��� ������ ���� ���� �����Ǿ��ִ� ��Ʈ�� ������ �̸����� Ȯ��
        
            CheckQuery = 1
        
        End If
    Next
    
End Function

'=====================================================================
'�� ����Ʈ ȣ�� ���� Ȯ��
Public Function CheckCategory()
    
    '---�� ����Ʈ ���� ù��° �� ���� üũ
    If �����_����.Value = "" Then
    
        CheckCategory = 1
        
    End If
    
End Function

'=====================================================================
'���� ���� ���� üũ
Public Function CheckFile(ByVal path_ As String) As Boolean
        
    CheckFile = (Dir(path_, vbDirectory) <> "") '---�Էµ� ��ο� �Էµ� ���ϸ��� �����ϴ��� Ȯ��
 
End Function

'=====================================================================
'���� ���� ���� üũ
Public Function CheckFileOpen(CheckFile As String) As Boolean

    Dim wb As Variant
    
    On Error Resume Next
    
    Set wb = Workbooks(CheckFile)
        
    If Not wb Is Nothing Then
    
        CheckFileOpen = True
        
    Else
    
        CheckFileOpen = False
        
    End If
    
    On Error GoTo 0
    
End Function

'=====================================================================
'������ �̸� üũ
Public Function CheckPresetName()
    
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
    CheckPresetName = "������" & preset_name_index
    
End Function

