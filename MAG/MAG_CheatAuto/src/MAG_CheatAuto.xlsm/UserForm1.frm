VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���� ���"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_FolderSearch_Click()
    
    Dim Selected As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "������ �����ϼ���"
        
        Selected = .Show
        If Selected = -1 Then
            Sheets("etc").Range("H2") = .SelectedItems(1)
        Else
            MsgBox "���õ� ������ �����ϴ�."
            End
        End If
        
    End With
End Sub
