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

'======================================================================================================
'[폴더 검색] 버튼 클릭 시 동작 이벤트


Private Sub Button_FolderSearch_Click()
    
    Dim Selected As Long '---선택한 파일 정보 저장 변수
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "폴더를 선택하세요"
        
        Selected = .Show '---파일 탐색기 열기
        
        '선택된 폴더가 있는 경우 동작
        If Selected = -1 Then
        
            Sheets("etc").Range("H2") = .SelectedItems(1)
        
        '선택된 폴더가 없는 경우 알림
        Else
        
            MsgBox "선택된 폴더가 없습니다."
            
        End If
    End With
    
End Sub

'======================================================================================================
'[닫기] 버튼 클릭 시 동작 이벤트


Private Sub CommandButton1_Click()
    
    '폼 종료
    Unload Me
    
End Sub
