VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "폴더 경로"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
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
        .title = "폴더를 선택하세요"
        
        Selected = .Show
        If Selected = -1 Then
            Sheets("etc").Range("H2") = .SelectedItems(1)
        Else
            MsgBox "선택된 폴더가 없습니다."
            End
        End If
        
    End With
End Sub
