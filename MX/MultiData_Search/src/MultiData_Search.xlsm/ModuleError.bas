Attribute VB_Name = "ModuleError"
Option Explicit

'=====================================================================
'��ũ�� : CheckUserData
'��� ��Ʈ : Main ��Ʈ
'���� : ����ڰ� �����͸� �Է��ߴ��� üũ�մϴ�.
'=====================================================================
Public Function CheckUserData() As Boolean
    
    '���� ��� �Է� üũ
    If ���ϰ��(1) = "" Then
        MsgBox "���� ��θ� �Է����ּ���."
        CheckUserData = True
        Exit Function
    
    '�˻��� �Է� üũ
    ElseIf �˻��� = "" Then
        MsgBox "�˻�� �Է����ּ���."
        CheckUserData = True
        Exit Function
        
    End If
    
    '���ϸ� �Է� üũ
    For i = 1 To ���ϸ�.count
    
        If ���ϸ�(i) = "" Then
            MsgBox "���ϸ��� �Է����ּ���."
            CheckUserData = True
            Exit Function
            
        End If
    Next
    
End Function

'=====================================================================
'��ũ�� : CheckFile
'���� : ����ڰ� �Է��� ������ ������ �����ϴ��� üũ�մϴ�.
'=====================================================================
Public Function CheckFile() As Boolean
    
    Dim strFile As String
    
    '�Էµ� ���� ������ŭ �ݺ�
    For j = 1 To ���ϸ�.count
        
        '���� �������� Ȯ��
        If InStr(���ϸ�(j), ".xl") = 0 Then
            MsgBox ���ϸ�(i) & "��(��) ���� ������ ������ �ƴմϴ�."
            CheckFile = True
            Exit Function
            
        End If
        
        strFile = ���ϰ��(j) & "\" & ���ϸ�(j)
        
        '�Էµ� ��ο� �Էµ� ���ϸ��� �����ϴ��� Ȯ��
        If Dir(strFile, vbDirectory) = "" Then
            MsgBox strFile & "��(��) �������� �ʴ� �����Դϴ�."
            CheckFile = True
            Exit Function
            
        End If
        
    Next
    
End Function

'=====================================================================
'��ũ�� : CheckSheet
'���� : ����ڰ� �Է��� ���� �� �Է��� ��Ʈ���� �����ϴ��� üũ�մϴ�.
'=====================================================================
Public Function CheckSheet(wb, strSheet) As Boolean

    For j = 1 To wb.Sheets.count
        
        '��Ʈ���� ��ġ�ϴ��� üũ
        If wb.Sheets(j).Name = strSheet Then
            Exit Function
            
        End If
    Next
        
    CheckSheet = True
    
End Function

'=====================================================================
'���� ���� ���� üũ
Public Function CheckFileOpen(CheckFile As String) As Boolean

    Dim wb1 As Variant
    
    On Error Resume Next
    
    Set wb1 = Workbooks(Dir(CheckFile))
        
    If Not wb1 Is Empty Then
    
        CheckFileOpen = True
        
    Else
    
        CheckFileOpen = False
        
    End If
    
    On Error GoTo 0
    
End Function
