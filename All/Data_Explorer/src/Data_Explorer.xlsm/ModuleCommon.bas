Attribute VB_Name = "ModuleCommon"
'=====================================================================
'기타 사용 모듈
'=====================================================================

'###전역 변수###

Public File_adr As String
Public File_name As String
Public preset As String
Public sheet_name As String

'###영역 전역 변수###

Public 파일경로 As Range
Public 파일명 As Range
Public 시트명 As Range
Public 프리셋명 As Range

Public 현재프리셋 As Range
Public 열목록 As Range
Public 열목록_시작 As Range
Public 열목록_끝 As Range

Public 검색어_시작 As Range
Public 검색키워드_시작 As Range
Public 검색키워드_끝 As Range
Public 고정행 As Range

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
'동일한 시트명 체크
Public Function CheckQuery()
    
    Dim sheet_count As Variant '---시트 개수 저장 변수
        
    sheet_count = ActiveWorkbook.Sheets.Count '---현재 파일에서 시트 개수 체크
    
    '---시트 개수만큼 반복
    For i = 1 To sheet_count
        
        If preset = ActiveWorkbook.Sheets(i).Name Then '---입력한 프리셋 명이 현재 생성되어있는 시트와 동일한 이름인지 확인
        
            CheckQuery = 1
        
        End If
    Next
    
End Function

'=====================================================================
'열 리스트 호출 여부 확인
Public Function CheckCategory()
    
    '---열 리스트 영역 첫번째 셀 공백 체크
    If 열목록_시작.Value = "" Then
    
        CheckCategory = 1
        
    End If
    
End Function

'=====================================================================
'주 사용 영역 지정
Public Sub SetRange()
    
    With Sheets("home")
                
        Set 파일경로 = .Range("C4") '---파일 경로
        Set 파일명 = .Range("C5") '---파일 이름
        Set 시트명 = .Range("C6") '---시트 목록
        Set 프리셋명 = .Range("C7") '---프리셋 이름
        
        Set 현재프리셋 = .Range("G4") '---현재 프리셋 이름
        Set 열목록_시작 = 현재프리셋.Offset(1, 0) '---열 리스트 처음 위치
        Set 열목록_끝 = 현재프리셋.End(xlDown) '---열 리스트 마지막 위치
        Set 열목록 = Range(열목록_시작, 열목록_끝) '---열 리스트 전체 영역
        
        Set 검색어_시작 = .Range("K4") '---검색 시작 영역
        Set 검색키워드_시작 = .Range("K5") '---선택된 열 시작 영역
        Set 고정행 = .Range("J8") '---행 고정 입력 영역
        
        '---선택된 열 끝 영역
        If 검색키워드_시작 = "" Then
            
            Set 검색키워드_끝 = 검색키워드_시작
            
        Else
        
            Set 검색키워드_끝 = 검색키워드_시작.Offset(0, -1).End(xlToRight)
            
        End If
        
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
    
    '--- sub 검색중인 내용 초기화
    ResetSearch
    
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

'=====================================================================
'파일 존재 여부 체크
Public Function CheckFile(ByVal path_ As String) As Boolean
        
    CheckFile = (Dir(path_, vbDirectory) <> "") '---입력된 경로에 입력된 파일명이 존재하는지 확인
 
End Function

'=====================================================================
'파일 오픈 상태 체크
Public Function CheckFileOpen(CheckFile As String) As Boolean

    Dim wb As Variant
    
    On Error Resume Next
    
    Set wb = Workbooks(CheckFile)
        
    If Not wb Is Nothing Then
    
        CheckFileOpen = True
        
    Else
    
        CheckFileOpen = False
        
    End If
    
    On Error GoTo 0
    
End Function

'=====================================================================
'프리셋 이름 체크
Public Function CheckPresetName()
    
    Dim preset_name_index As Variant '---프리셋명 순서
    Dim check As Boolean '---동일한 프리셋명 체크
    
    preset_name_index = 1
    
    With Range("preset_list")
        Do
            For i = 2 To .Cells.Count
                
                '---동일한 프리셋 이름이 존재하는 경우 처리
                If StrComp(.Cells(i).Value, "프리셋" & preset_name_index) = 0 Then
                
                    preset_name_index = preset_name_index + 1
                    check = True
                    Exit For
                    
                End If
                
                check = False
            Next
            
            '---동일한 프리셋 이름이 없으면 종료
            If check = False Then
            
                Exit Do
                
            End If
        Loop
    End With
    
    '---프리셋 이름 반환
    CheckPresetName = "프리셋" & preset_name_index
    
End Function
