VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "폴더 경로"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
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
    Dim path As String
    Dim row As Integer
    
    Dim strFolderList As String
    Dim strSpl As Variant
    Dim temp As Integer
    Dim strLatest As String
    
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
    
    path = Sheets("etc").Range("H2").Value & "\"
    
    Sheets("etc").Columns("E:E").Clear
    
    row = 1
    
    '선택된 경로가 MAG 데이터 폴더가 아닌 경우 동작
    If InStr(path, "TFD_") = 0 Then
        
        strFolderList = Dir(path, vbDirectory)
        
        Do While strFolderList <> ""
            
            '폴더명이 . .. 인 경우 제외
            If strFolderList <> "." And strFolderList <> ".." And InStr(strFolderList, "TFD_") > 0 Then
            
                '폴더인지 확인
                If (GetAttr(path & strFolderList) And vbDirectory) = vbDirectory Then
                
                    Sheets("etc").Cells(row, 5).Value = strFolderList
                    row = row + 1
                        
                End If
                
            End If
            strFolderList = Dir
        Loop
        
        '검색된 폴더가 없는 경우 알림 처리
        If row = 1 Then
            MsgBox "선택된 경로에 MAG 데이터 폴더가 존재하지 않습니다."
        
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
