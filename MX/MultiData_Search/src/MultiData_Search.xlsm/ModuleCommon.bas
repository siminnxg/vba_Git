Attribute VB_Name = "ModuleCommon"
Option Explicit

'###전역 변수###

'영역 지정
Public 파일경로 As Range '---파일 경로 영역
Public 파일명, 파일명1 As Range '---파일 이름 영역
Public 시트명 As Range '---파일 시트 영역
Public 검색어 As Range '---검색어 입력 영역
Public 검색결과 As Range '---검색된 결과 영역
Public 오브젝트 As Range

'변수 지정
Public i, j, k As Variant '---반복문 사용 변수
Public dateTime As Date '---시간 체크용 변수
Public rngTemp As Range '---임의 영역 지정 변수

'=====================================================================
'영역 지정
Public Sub SetRange()
    
    With Sheets("Main")
    
        Set 파일경로 = .Range("B5") '---파일 경로 영역 지정
        
        '파일 이름 영역 지정
        Set 파일명 = .Range("B7")
        
        '파일 이름 다중 입력 시 처리
        If 파일명.Offset(1, 0) <> "" Then
            
            Set 파일명 = Range(파일명, 파일명.Offset(-1, 0).End(xlDown))
            
        End If
        
        Set 시트명 = .Range("C7") '---시트명 영역 지정
    
        Set 검색어 = .Range("B19") '---검색 값 영역 지정
        
        Set 검색결과 = .Range("B22") '---검색 결과 표시 영역 지정
        
    End With
    
    With Sheets("etc")
        
        Set 오브젝트 = .Range("A1") '---호출된 파일 리스트 저장 위치
        
    End With
End Sub

'=====================================================================
'화면 업데이트 중지 (동작 속도 증가)
Public Sub UpdateStart()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub

'=====================================================================
'화면 업데이트 원복
Public Sub UpdateEnd()
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub
