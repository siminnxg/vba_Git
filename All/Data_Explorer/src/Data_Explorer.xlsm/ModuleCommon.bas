Attribute VB_Name = "ModuleCommon"
'=====================================================================
'기타 사용 모듈
'=====================================================================

'###전역 변수###

Public File_adr As String
Public File_name As String
Public preset As String
Public sheet_name As String

Public i, j, k As Variant '---반복문 사용 변수

'###영역 전역 변수###

Public 파일경로 As Range
Public 파일명 As Range
Public 시트명 As Range
Public 프리셋명 As Range

Public 현재프리셋 As Range
Public 열목록, 열목록_시작, 열목록_끝 As Range

Public 검색어_시작 As Range
Public 검색키워드, 검색키워드_시작, 검색키워드_끝 As Range
Public 고정행 As Range
Public 틀고정 As Range

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

'=====================================================================
'주 사용 영역 지정
Public Sub SetRange()
    
    With Sheets("Home")
                
        Set 파일경로 = .Range("C4") '---파일 경로
        Set 파일명 = .Range("C5") '---파일 이름
        Set 시트명 = .Range("C6") '---시트 목록
        Set 프리셋명 = .Range("C7") '---프리셋 이름
        
    End With
    
    With Sheets("Search")
    
        Set 현재프리셋 = .Range("B4") '---현재 프리셋 이름
        Set 열목록_시작 = 현재프리셋.Offset(1, 0) '---열 리스트 처음 위치
        Set 열목록_끝 = 현재프리셋.End(xlDown) '---열 리스트 마지막 위치
        Set 열목록 = Range(열목록_시작, 열목록_끝) '---열 리스트 전체 영역
        
        Set 검색어_시작 = .Range("F4") '---검색 시작 영역
        Set 검색키워드_시작 = 검색어_시작.Offset(1, 0) '---선택된 열 시작 영역
        Set 고정행 = .Range("E8") '---행 고정 입력 영역
        Set 틀고정 = .Range("E10")
        
        '---선택된 열 끝 영역
        If 검색키워드_시작 = "" Then
            Set 검색키워드_끝 = 검색키워드_시작
        Else
            Set 검색키워드_끝 = 검색키워드_시작.Offset(0, -1).End(xlToRight)
        End If
        
        Set 검색키워드 = Range(검색키워드_시작, 검색키워드_끝)
        
    End With
    
    '---sub 전역 변수에 값 저장
    Call LoadFileInfo

End Sub

'=====================================================================
'사용자가 입력한 정보 전역 변수에 저장
Public Sub LoadFileInfo()
    
    File_adr = 파일경로.Value '---파일 경로 저장
    File_name = 파일명.Value '---파일 이름 저장
    sheet_name = 시트명.Value '----시트 이름 저장
    preset = 프리셋명.Value '---프리셋 이름 저장
    
End Sub

'=====================================================================
'현재 검색중인 내용, 카테고리 초기화
Public Sub ClearHomeData()
    
    '---현재 불러온 데이터가 없으면 종료
    If 현재프리셋.Value = Empty Then
        
        Exit Sub
    
    End If
    
    Range(검색어_시작, 검색키워드_끝).Clear
    
    Range("DATA").ClearContents '---열 선택 영역 초기화
    
    Range("DATA").FormatConditions.Delete
    
    '--- sub 열 리스트 선택 초기화
    Call ResetCategory
    
    Range("notice").ClearContents
    
    
End Sub
      
'=====================================================================
'연결 한번에 삭제
Public Sub DeleteConnect()

    Dim conn As Object
    Dim connName As String
    
    '---모든 연결을 순회
    For Each conn In ActiveWorkbook.Connections
        connName = conn.Name '---연결 이름 가져오기
        
        '---연결로 시작하는 이름의 연결 제거
        If connName Like "연결*" Then
        
            conn.Delete
            
        End If
    Next conn
    
End Sub
