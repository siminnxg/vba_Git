Attribute VB_Name = "ModuleCommon"
Option Explicit

Public 파일경로 As Range

Public 검색어 As Range

Public 검색목록 As Range
Public 검색목록_시작 As Range
Public 검색목록_끝 As Range

Public 키목록 As Range
Public 키목록_시작 As Range
Public 키목록_끝 As Range

Public 치트키 As Range

Public 타입 As ListObject

Public i, j, k As Variant '---반복문 사용 변수
Public cell As Range

Public Sub SetRange()

    With Sheets("Main")
        
        Set 검색어 = .Range("B6")
        
        '키목록 영역 지정
        Set 키목록_시작 = .Range("B9")
        If IsError(키목록_시작.Value) Then
            Set 키목록_끝 = 키목록_시작
        Else
            Set 키목록_끝 = 키목록_시작.Offset(-1, 0).End(xlDown)
        End If
        Set 키목록 = Range(키목록_시작, 키목록_끝)
        
        '파일경로 영역 지정
        Set 파일경로 = .Range("B3")
        
        '검색목록 영역 지정
        Set 검색목록_시작 = .Range("E3")
        If 검색목록_시작.Value = "" Then
            Set 검색목록_끝 = 검색목록_시작
        Else
            Set 검색목록_끝 = 검색목록_시작.Offset(-1, 0).End(xlDown)
        End If
        Set 검색목록 = Range(검색목록_시작, 검색목록_끝)
        
        '치트키 영역 지정
        Set 치트키 = .Range("K3")
        
    End With
    
    With Sheets("etc")
        
        Set 타입 = .ListObjects(1)
        
    End With
    
End Sub

Public Sub UpdateStart()
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


End Sub


Public Sub UpdateEnd()

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
