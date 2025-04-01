Attribute VB_Name = "ModuleCommon"
'=====================================================================
'��Ÿ ��� ���
'=====================================================================

'###���� ����###

Public File_adr As String
Public File_name As String
Public preset As String
Public sheet_name As String

Public i, j, k As Variant '---�ݺ��� ��� ����

'###���� ���� ����###

Public ���ϰ�� As Range
Public ���ϸ� As Range
Public ��Ʈ�� As Range
Public �����¸� As Range

Public ���������� As Range
Public �����, �����_����, �����_�� As Range

Public �˻���_���� As Range
Public �˻�Ű����, �˻�Ű����_����, �˻�Ű����_�� As Range
Public ������ As Range
Public Ʋ���� As Range

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
'�� ��� ���� ����
Public Sub SetRange()
    
    With Sheets("Home")
                
        Set ���ϰ�� = .Range("C4") '---���� ���
        Set ���ϸ� = .Range("C5") '---���� �̸�
        Set ��Ʈ�� = .Range("C6") '---��Ʈ ���
        Set �����¸� = .Range("C7") '---������ �̸�
        
    End With
    
    With Sheets("Search")
    
        Set ���������� = .Range("B4") '---���� ������ �̸�
        Set �����_���� = ����������.Offset(1, 0) '---�� ����Ʈ ó�� ��ġ
        Set �����_�� = ����������.End(xlDown) '---�� ����Ʈ ������ ��ġ
        Set ����� = Range(�����_����, �����_��) '---�� ����Ʈ ��ü ����
        
        Set �˻���_���� = .Range("F4") '---�˻� ���� ����
        Set �˻�Ű����_���� = �˻���_����.Offset(1, 0) '---���õ� �� ���� ����
        Set ������ = .Range("E8") '---�� ���� �Է� ����
        Set Ʋ���� = .Range("E10")
        
        '---���õ� �� �� ����
        If �˻�Ű����_���� = "" Then
            Set �˻�Ű����_�� = �˻�Ű����_����
        Else
            Set �˻�Ű����_�� = �˻�Ű����_����.Offset(0, -1).End(xlToRight)
        End If
        
        Set �˻�Ű���� = Range(�˻�Ű����_����, �˻�Ű����_��)
        
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
