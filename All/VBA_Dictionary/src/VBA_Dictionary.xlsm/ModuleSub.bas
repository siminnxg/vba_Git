Attribute VB_Name = "ModuleSub"
Option Explicit

Sub ��������()

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
