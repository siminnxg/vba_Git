Attribute VB_Name = "ModuleCommon"
Option Explicit

'###################################################
'공용 사용 변수, 모듈
'###################################################


'Search 영역
Public 검색어 As Range
Public 키목록 As Range
Public 키목록_시작 As Range
Public 키목록_끝 As Range

'치트키 설정 영역
Public 검색목록 As Range
Public 검색목록_시작 As Range
Public 검색목록_끝 As Range
Public 검색옵션_시작 As Range
Public 검색옵션_스텟 As Range

'Cheat List 영역
Public 치트키 As Range
Public 치트키_시작 As Range
Public 치트키_끝 As Range
Public 프리셋 As Range
Public 프리셋_끝 As Range

'etc 시트 영역
Public 파일경로 As Range
Public 타입 As ListObject

Public rngCheat1 As Range '---아이템 생성 치트키 영역
Public rngCheat2 As Range '---랜덤 옵션 아이템 생성 치트키 영역

Public i, j, k As Variant
Public cnt As Variant
Public cell, cell2 As Range


Public Sub SetRange()

    With ThisWorkbook.Sheets("Main")
        
        '치트키1, 2 영역 지정
        Set rngCheat1 = .Range("E:E,H:J").Columns
        Set rngCheat2 = .Range("K:L,P:P,S:T,W:W").Columns
        
        Set 검색어 = .Range("B7")
        
        '키목록 영역 지정
        Set 키목록_시작 = .Range("B10")
        If IsError(키목록_시작.Value) Then
            Set 키목록_끝 = 키목록_시작
        Else
            Set 키목록_끝 = 키목록_시작.Offset(-1, 0).End(xlDown)
        End If
        Set 키목록 = Range(키목록_시작, 키목록_끝)
        
        '검색목록 영역 지정
        'Cheat1 / Cheat2 구분
        If rngCheat2.Hidden = True Then
            Set 검색목록_시작 = .Range("E7")
        Else
            Set 검색목록_시작 = .Range("L7")
            Set 검색옵션_시작 = 검색목록_시작.Offset(0, 3)
            Set 검색옵션_스텟 = .Range("T5")
        End If
        
        If 검색목록_시작.Value = "" Then
            Set 검색목록_끝 = 검색목록_시작
        Else
            Set 검색목록_끝 = 검색목록_시작.Offset(-1, 0).End(xlDown)
        End If
        Set 검색목록 = Range(검색목록_시작, 검색목록_끝)
        
        '치트키 영역 지정
        Set 치트키_시작 = .Range("X7")
        If IsEmpty(치트키_시작) Then
            Set 치트키_끝 = 치트키_시작
        Else
            Set 치트키_끝 = 치트키_시작.Offset(-1, 0).End(xlDown)
        End If
        Set 치트키 = Range(치트키_시작, 치트키_끝)
        
        Set 프리셋 = 치트키_시작.Offset(0, 2)
        Set 프리셋_끝 = 프리셋.Offset(1, 0).End(xlDown)
    End With
    
    With ThisWorkbook.Sheets("etc")
        
        Set 타입 = .ListObjects(1)
        
        Set 파일경로 = .Range("H2")
        
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
