Attribute VB_Name = "ModuleCommon"
'=====================================================================
'��Ÿ ��� ���
'=====================================================================

'###���� ����###

Public File_adr As String
Public File_name As String
Public preset As String
Public sheet_name As String

'###���� ���� ����###

Public ���ϰ�� As Range
Public ���ϸ� As Range
Public ��Ʈ�� As Range
Public �����¸� As Range

Public ���������� As Range
Public ����� As Range
Public �����_���� As Range
Public �����_�� As Range

Public �˻���_���� As Range
Public �˻�Ű����_���� As Range
Public �˻�Ű����_�� As Range
Public ������ As Range

'=====================================================================
'ȭ�� ������Ʈ ���� (���� �ӵ� ����)
Public Sub UpdateStart()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub

'=====================================================================
'ȭ�� ������Ʈ ����
Public Sub UpdateEnd()
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

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
'�� ��� ���� ����
Public Sub SetRange()
    
    With Sheets("home")
                
        Set ���ϰ�� = .Range("C4") '---���� ���
        Set ���ϸ� = .Range("C5") '---���� �̸�
        Set ��Ʈ�� = .Range("C6") '---��Ʈ ���
        Set �����¸� = .Range("C7") '---������ �̸�
        
        Set ���������� = .Range("G4") '---���� ������ �̸�
        Set �����_���� = ����������.Offset(1, 0) '---�� ����Ʈ ó�� ��ġ
        Set �����_�� = ����������.End(xlDown) '---�� ����Ʈ ������ ��ġ
        Set ����� = Range(�����_����, �����_��) '---�� ����Ʈ ��ü ����
        
        Set �˻���_���� = .Range("K4") '---�˻� ���� ����
        Set �˻�Ű����_���� = .Range("K5") '---���õ� �� ���� ����
        Set ������ = .Range("J8") '---�� ���� �Է� ����
        
        '---���õ� �� �� ����
        If �˻�Ű����_���� = "" Then
            
            Set �˻�Ű����_�� = �˻�Ű����_����
            
        Else
        
            Set �˻�Ű����_�� = �˻�Ű����_����.Offset(0, -1).End(xlToRight)
            
        End If
        
    End With
    
    '---sub ���� ������ �� ����
    Call LoadFileInfo

End Sub

'=====================================================================
'����ڰ� �Է��� ���� ���� ������ ����
Public Sub LoadFileInfo()
    
    File_adr = ���ϰ��.Value '---���� ��� ����
    File_name = ���ϸ�.Value '---���� �̸� ����
    sheet_name = ��Ʈ��.Value '----��Ʈ �̸� ����
    preset = �����¸�.Value '---������ �̸� ����
    
End Sub

'=====================================================================
'���� �˻����� ����, ī�װ� �ʱ�ȭ
Public Sub ClearHomeData()
    
    '---���� �ҷ��� �����Ͱ� ������ ����
    If ����������.Value = Empty Then
        
        Exit Sub
    
    End If
    
    '--- sub �˻����� ���� �ʱ�ȭ
    ResetSearch
    
    Range(�˻���_����, �˻�Ű����_��).Clear
    
    Range("DATA").ClearContents '---�� ���� ���� �ʱ�ȭ
    
    Range("DATA").FormatConditions.Delete
    
    '--- sub �� ����Ʈ ���� �ʱ�ȭ
    Call ResetCategory
    
    Range("notice").ClearContents
    
    
End Sub
      
'=====================================================================
'���� �ѹ��� ����
Public Sub DeleteConnect()

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
