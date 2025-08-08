VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "TOOL"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##########################################################
'설정된 폴더 경로 확인 및 변경하기
'##########################################################


Private Sub Btn_FileAdr_Click()
    
    If MsgBox("현재 설정된 경로는 '" & ThisWorkbook.CustomDocumentProperties("폴더경로").Value & "' 입니다." & vbCrLf & vbCrLf & _
        "경로를 변경하시겠습니까?", vbQuestion + vbOKCancel) _
        = 1 Then
    
        Call SearchFolder
        Unload Me
    
    End If
    
End Sub


'##########################################################
'Result 영역, 결과 표시 영역에 함수 전체 입력
'##########################################################


Private Sub Btn_ShtRefresh_Click()
    
    Dim row As Variant
    
    Call Unprotect
    
    row = Range("사용자입력")(1).row
    
    Range("결과").Formula2 = "=결과확인($B" & row & ":$E" & row & ",$H" & row & "#)"
    
    Range("결과표시").Formula2 = "=FILTER(GachaElement,(GachaElement[[long_2]]=ID추출($C" & row & "))*(GachaElement[[long_1]]=Main!$B" & row & "), """")"
    
    Sheets("Main").Protect
    
    MsgBox "수식 새로고침이 완료되었습니다."
    
    Unload Me
    
End Sub


'##########################################################
'메모 영역 행 추가
'##########################################################

Private Sub Btn_AddRow_Click()
    
    Call Unprotect
    
    With Sheets("Main")
        .Rows(Range("MEMO").row).Insert
        .Protect
    End With
    
End Sub


'##########################################################
'메모 영역 행 제거
'##########################################################


Private Sub Btn_DelRow_Click()
    
    Dim row As Variant
    
    row = Range("MEMO").row
    
    If row < 6 Then
        MsgBox "제거할 메모 공간이 부족합니다."
        Exit Sub
    End If
    
    Call Unprotect
        
    With Sheets("Main")
        .Rows(row - 1).Delete
        .Protect
    End With
    
End Sub
