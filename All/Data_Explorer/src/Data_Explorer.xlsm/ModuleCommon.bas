Attribute VB_Name = "ModuleCommon"
'=====================================================================
'기타 사용 모듈
'=====================================================================

'###전역 변수###

Public File_name As String
Public File_adr As String
Public preset As String
Public sheet_name As String

'###영역 전역 변수###

Public user_file_adr As Range
Public user_file_name As Range
Public user_file_sheet As Range
Public user_file_preset As Range

Public act_sheet_name As Range
Public act_category_list As Range
Public act_category_start As Range
Public act_category_end As Range

Public search_user_start As Range
Public search_category_start As Range
Public search_FixRow As Range
Public search_category_end As Range

Public etc_preset As Range

'=====================================================================
'화면 업데이트 중지 (동작 속도 증가)
Public Sub update_start()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub

'=====================================================================
'화면 업데이트 원복
Public Sub update_end()
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

'=====================================================================
'동일한 시트명 체크
Public Function query_check()
    
    Dim sheet_count As Variant '---시트 개수 저장 변수
        
    sheet_count = ActiveWorkbook.Sheets.Count '---현재 파일에서 시트 개수 체크
    
    '---시트 개수만큼 반복
    For i = 1 To sheet_count
        
        If preset = ActiveWorkbook.Sheets(i).Name Then '---입력한 프리셋 명이 현재 생성되어있는 시트와 동일한 이름인지 확인
        
            query_check = 1
        
        End If
    Next
    
End Function

'=====================================================================
'열 리스트 호출 여부 확인
Public Function category_check()
    
    '---열 리스트 영역 첫번째 셀 공백 체크
    If act_category_start.Value = "" Then
    
        category_check = 1
        
    End If
    
End Function

'=====================================================================
'주 사용 영역 지정
Public Sub range_set()
    
    With Sheets("home")
                
        Set user_file_adr = .Range("C4") '---파일 경로
        Set user_file_name = .Range("C5") '---파일 이름
        Set user_file_sheet = .Range("C6") '---시트 목록
        Set user_file_preset = .Range("C7") '---프리셋 이름
        
        Set act_sheet_name = .Range("G4") '---현재 프리셋 이름
        Set act_category_start = act_sheet_name.Offset(1, 0) '---열 리스트 처음 위치
        Set act_category_end = act_sheet_name.End(xlDown) '---열 리스트 마지막 위치
        Set act_category_list = Range(act_category_start, act_category_end) '---열 리스트 전체 영역
        
        Set search_user_start = .Range("K4") '---검색 시작 영역
        Set search_category_start = .Range("K5") '---선택된 열 시작 영역
        Set search_FixRow = .Range("J8") '---행 고정 입력 영역
        
        '---선택된 열 끝 영역
        If search_category_start = "" Then
            
            Set search_category_end = search_category_start
            
        Else
        
            Set search_category_end = search_category_start.Offset(0, -1).End(xlToRight)
            
        End If
        
    End With
    
    With Sheets("etc")
    
        Set etc_preset = .Range("H2") '---프리셋 선택 영역
    
    End With
    
    '---sub 전역 변수에 값 저장
    Call file_info_load

End Sub

'=====================================================================
'사용자가 입력한 정보 전역 변수에 저장
Public Sub file_info_load()
    
    File_adr = user_file_adr.Value '---파일 경로 저장
    File_name = user_file_name.Value '---파일 이름 저장
    sheet_name = user_file_sheet.Value '----시트 이름 저장
    preset = user_file_preset.Value '---프리셋 이름 저장
    
End Sub

'=====================================================================
'현재 검색중인 내용, 카테고리 초기화
Public Sub home_data_clear()
    
    '---현재 불러온 데이터가 없으면 종료
    If act_sheet_name.Value = Empty Then
        
        Exit Sub
    
    End If
    
    '--- sub 검색중인 내용 초기화
    search_reset
    
    Range(search_user_start, search_category_end).Clear
    
    Range("DATA").ClearContents '---열 선택 영역 초기화
    
    Range("DATA").FormatConditions.Delete
    
    '--- sub 열 리스트 선택 초기화
    Call category_reset
    
    Range("notice").ClearContents
    
    
End Sub
      
'=====================================================================
'연결 한번에 삭제
Public Sub connect_delete()

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
Public Function FileExists(ByVal path_ As String) As Boolean
        
    FileExists = (Dir(path_, vbDirectory) <> "") '---입력된 경로에 입력된 파일명이 존재하는지 확인
 
End Function

'=====================================================================
'파일 오픈 상태 체크
Public Function IsCheckOpen(CheckFile As String) As Boolean

    Dim wb As Variant
    
    On Error Resume Next
    
    Set wb = Workbooks(CheckFile)
        
    If Not wb Is Nothing Then
    
        IsCheckOpen = True
        
    Else
    
        IsCheckOpen = False
        
    End If
    
    On Error GoTo 0
    
End Function

'=====================================================================
'프리셋 이름 체크
Public Function preset_name_check()
    
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
    preset_name_check = "프리셋" & preset_name_index
    
End Function
