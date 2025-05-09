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

Private Sub Button_FolderSearch_Click()
    
    Dim Selected As Long
    Dim path As String
    Dim row As Integer
    
    Dim strFolderList As String
    Dim strSpl As Variant
    Dim temp As Integer
    Dim strLatest As String
    
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
    
    path = Sheets("etc").Range("H2").Value & "\"
    
    Sheets("etc").Columns("E:E").Clear
    
    row = 1
    
    '���õ� ��ΰ� MAG ������ ������ �ƴ� ��� ����
    If InStr(path, "TFD_") = 0 Then
        
        strFolderList = Dir(path, vbDirectory)
        
        Do While strFolderList <> ""
            
            '�������� . .. �� ��� ����
            If strFolderList <> "." And strFolderList <> ".." And InStr(strFolderList, "TFD_") > 0 Then
            
                '�������� Ȯ��
                If (GetAttr(path & strFolderList) And vbDirectory) = vbDirectory Then
                
                    Sheets("etc").Cells(row, 5).Value = strFolderList
                    row = row + 1
                        
                End If
                
            End If
            strFolderList = Dir
        Loop
        
        '�˻��� ������ ���� ��� �˸� ó��
        If row = 1 Then
            MsgBox "���õ� ��ο� MAG ������ ������ �������� �ʽ��ϴ�."
        
        Else
            temp = 0
            
            For i = 1 To row
            
                With Sheets("etc").Cells(i, 5)
                    
                    strSpl = Split(.Value, "_CL")
                            
                    If UBound(strSpl) = 1 Then
                        
                        If Val(strSpl(1)) > temp Then
                        
                            temp = Val(strSpl(1))
                            strLatest = .Value
                            
                        End If
                        
                    End If
                    
                End With
                
            Next
        
            Sheets("etc").Range("H2") = path & strLatest
        End If
    End If
    
End Sub
