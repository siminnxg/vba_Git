VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���� ���"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'======================================================================================================
'[���� �˻�] ��ư Ŭ�� �� ���� �̺�Ʈ


Private Sub Button_FolderSearch_Click()
    
    Dim Selected As Long '---������ ���� ���� ���� ����
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "������ �����ϼ���"
        
        Selected = .Show '---���� Ž���� ����
        
        '���õ� ������ �ִ� ��� ����
        If Selected = -1 Then
        
            Sheets("etc").Range("H2") = .SelectedItems(1)
        
        '���õ� ������ ���� ��� �˸�
        Else
        
            MsgBox "���õ� ������ �����ϴ�."
            
        End If
    End With
    
End Sub

'======================================================================================================
'[�ݱ�] ��ư Ŭ�� �� ���� �̺�Ʈ


Private Sub CommandButton1_Click()
    
    '�� ����
    Unload Me
    
End Sub
