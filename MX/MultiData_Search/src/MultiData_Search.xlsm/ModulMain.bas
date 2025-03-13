Attribute VB_Name = "ModulMain"
Option Explicit

'=====================================================================
'매크로 : SearchData
'대상 시트 : Main 시트
'동작 : 사용자가 입력한 데이터 파일들에서 검색어를 찾아 해당 행의 데이터를 모두 가져옵니다.
'=====================================================================
Public Sub SearchData()
    
    '호출 파일 관련 변수
    Dim Obj As Object
    Dim Wb As Workbook
    Dim WS As Worksheet
    
    Dim strFile As String '---파일 경로 & 이름
    Dim strSheet As String '---시트명
    
    '호출한 파일 내에서 영역 지정
    Dim rngWS_Search As Range '---검색 위치 저장
    Dim rngWS_Result As Range '---검색 결과 저장
    Dim rngWS_Col As Variant '---검색중인 시트의 열 개수
    
    Dim strSearchStart As String '---첫번째 검색어 주소 저장
    Dim varResultCount As Variant '---검색된 개수
    
    
    '###동작 시작###
    'On Error Resume Next
        
    Call UpdateStart '---화면 업데이트 중지
    Call SetRange '---주 사용 영역 지정
    
    '사용자 입력 데이터 없는 경우 처리
    If CheckUserData = True Then
        
        GoTo exit_sub
    
    '경로에 파일 없는 경우 처리
    ElseIf CheckFile() = True Then
                
        GoTo exit_sub
        
    End If
    
    varResultCount = 0
    
    '검색할 개수 확인
    For i = 1 To 파일명.count
        
        strFile = 파일경로.Value & "\" & 파일명(i)
        strSheet = 파일명(i).Offset(0, 1).Value
        
        '파일 정보 etc 시트 '오브젝트' 영역에 저장
        If 오브젝트 = "" Then
            Range("오브젝트")(i) = strFile
        
        Else
            j = Application.WorksheetFunction.CountIf(Range("오브젝트"), strFile)
            
            If j = 0 Then
                오브젝트.Offset(Range("오브젝트").count, 0) = strFile
                ThisWorkbook.Names("오브젝트").RefersTo = Range(Range("오브젝트"), 오브젝트.End(xlDown)) '---파일 개수가 2개 이상인 경우 '오브젝트' 영역 재지정
            End If
            
        End If
    
        '파일 호출
        Set Obj = GetObject(strFile)
        Set Wb = Workbooks(Dir(strFile))
                       
        '파일에 해당 시트 없는 경우 종료
        If CheckSheet(Wb, strSheet) = True Then
            MsgBox 파일명(i) & " 파일에 " & strSheet & " 시트가 존재하지 않습니다."
            
            Obj.Close '---호출된 파일 닫기
            GoTo exit_sub
            
        End If
        
        Set WS = Wb.Sheets(strSheet)
                
        '호출된 파일 내에서 검색 개수 체크
        If 검색어.Offset(0, 1) = "포함" Then
            varResultCount = varResultCount + Application.WorksheetFunction.CountIf(WS.UsedRange, "*" & 검색어 & "*") '---검색 옵션 포함
            
        Else
            varResultCount = varResultCount + Application.WorksheetFunction.CountIf(WS.UsedRange, 검색어) '---검색 옵션 일치
            
        End If
    Next
    
    '검색 결과 1만개 이상 시 종료
    If varResultCount > 10000 Then
        
        MsgBox "검색된 결과가 " & Format(varResultCount, "0,000") & "개 입니다. " & _
                vbCrLf & "데이터가 많아 조회에 많은 시간이 소요됩니다." & _
                vbCrLf & vbCrLf & "자세한 검색어를 입력해주세요."
                
            GoTo exit_sub
    
    '검색된 결과가 없는 경우 종료
    ElseIf varResultCount = 0 Then
        
        MsgBox "검색 결과가 없습니다."
        GoTo exit_sub
        
    End If
    
    '검색 결과 표시되는 'DATA' 영역 초기화
    Range("DATA").Clear
    ThisWorkbook.Names("DATA").RefersTo = 검색결과
    
    '입력된 파일 개수만큼 반복
    For i = 1 To 파일명.count
    
        strFile = 파일경로.Value & "\" & 파일명(i)
        strSheet = 파일명(i).Offset(0, 1).Value
        
        varResultCount = 0 '---검색 개수 초기화
        
        '검색할 파일 설정
        Set Wb = Workbooks(Dir(strFile))
        Set WS = Wb.Sheets(strSheet)
        
        '검색
        If 검색어.Offset(0, 1) = "포함" Then
            
            Set rngWS_Search = WS.UsedRange.Find(what:=검색어, lookat:=xlPart)
            
        Else
            
            Set rngWS_Search = WS.UsedRange.Find(what:=검색어, lookat:=xlWhole)
            
        End If
        
        '검색된 값이 존재하는 경우
        If Not rngWS_Search Is Nothing Then
                        
            strSearchStart = rngWS_Search.Address '---처음 검색한 위치 저장

            rngWS_Col = WS.UsedRange.Columns.count '---검색할 파일에서 사용중인 열 개수 체크
            
            Set rngWS_Result = WS.UsedRange.Rows(2) '---머릿글 행 설정 (MX는 2행 고정)
            
            Set rngWS_Result = Union(rngWS_Result, WS.UsedRange.Rows(rngWS_Search.Row)) '---검색된 행을 변수에 추가
            
            '무한 루프로 지속 검색
            Do
            
                Set rngWS_Search = WS.UsedRange.FindNext(rngWS_Search) '---검색

                Set rngWS_Result = Union(rngWS_Result, WS.UsedRange.Rows(rngWS_Search.Row)) '---검색된 행을 변수에 추가
                    
            Loop While Not rngWS_Search Is Nothing And strSearchStart <> rngWS_Search.Address '---검색 내용이 없거나 첫번째 주소로 돌아온 경우 종료
            
            '검색 결과 붙여넣기
            검색결과 = 파일명(i) '---첫번째 열에 검색된 파일명 표시
            rngWS_Result.Copy Destination:=검색결과.Offset(0, 1) '---서식 포함 붙여넣기
            
            '검색 시 행단위로 검색되어 동일한 행에 같은 검색 결과가 존재하는 경우 검색 개수 -1
            For Each rngTemp In rngWS_Result.Areas
                
                varResultCount = varResultCount + rngTemp.Rows.count
                
            Next rngTemp
        
            ThisWorkbook.Names("DATA").RefersTo = Range(Range("DATA"), 검색결과.Offset(varResultCount, rngWS_Col)) '---'DATA' 영역 재지정
            
            Set 검색결과 = 검색결과.Offset(varResultCount + 1, 0) '---'검색결과' 영역 재지정
            
            Application.Goto reference:=검색어.Offset(-2, -1), Scroll:=True  ' 원하는 셀로 이동 후 스크롤
            
        End If
    Next

'종료 처리
exit_sub:
    Call UpdateEnd

End Sub

'=====================================================================
'매크로 : CloseFile
'대상 시트 : etc 시트
'동작 : 사용자가 입력한 데이터 파일들에서 검색어를 찾아 해당 행의 데이터를 모두 가져옵니다.
'=====================================================================
Public Sub CloseFile()
    
    Dim Wb As Workbook
    Dim count As Variant
    
    On Error Resume Next
    
    '파일 미호출 상태인 경우 종료
    If Range("오브젝트")(1) = "" Then
        
        'MsgBox "닫을 파일이 존재하지 않습니다."
        Exit Sub
        
    End If
        
    Call SetRange '---주 사용 영역 지정
    
    '호출된 파일들 닫기
    For i = 1 To Range("오브젝트").count
        
        Set Wb = Workbooks(Dir(Range("오브젝트")(i)))
        Wb.Close
        count = 1
        
    Next
    
    If count = 1 Then
            
        ' '오브젝트' 영역 데이터 초기화
        Range(Range("오브젝트"), Range("오브젝트").Offset(0, 2)).Clear
        ThisWorkbook.Names("오브젝트").RefersTo = 오브젝트
    
    Else
        'MsgBox "닫을 파일이 존재하지 않습니다."
        
    End If
    
End Sub

'=====================================================================
'매크로 : OpenFile
'동작 : GetObject로 호출한 파일들을 모두 화면에 띄워줍니다.
'=====================================================================
Public Sub OpenFile()

    Dim Wb As Workbook
    
    On Error Resume Next
    
    If Range("오브젝트").Cells(1) = "" Then
        Exit Sub
    End If
    
    Call UpdateStart
    
    For i = 1 To Range("오브젝트").count
        Set Wb = Workbooks(Dir(Range("오브젝트").Cells(i)))
        
        Wb.IsAddin = True
        Wb.IsAddin = False
        ThisWorkbook.Activate
        
    Next
    
    Call UpdateEnd
End Sub

'=====================================================================
'매크로 : ClearSearch
'대상 시트 : Main 시트
'동작 : 검색 결과 영역을 초기화 합니다.
'=====================================================================
Public Sub ClearSearch()
    
    Call SetRange '---주 사용 영역 지정
    
    '검색 결과 표시되는 'DATA' 영역 초기화
    Range("DATA").Clear
    ThisWorkbook.Names("DATA").RefersTo = 검색결과
    
    검색어.ClearContents '--- 검색어 초기화
    
End Sub

'=====================================================================
'매크로 : SearchFile
'대상 시트 : Main 시트
'동작 : 파일을 검색하여 경로와 파일명을 조회합니다.
'=====================================================================
Public Sub SearchFile()
    
    Dim varFileNum As Variant
    
    Call SetRange '---주 사용 영역 지정
    
    '파일 경로가 입력되어 있으면 해당 경로로 지정
    '(잘못된 경로 입력 시 자동으로 무시됨)
    If 파일경로 <> "" Then
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = 파일경로
         
    End If
    
    '파일 탐색기 오픈
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "엑셀파일", "*.xls; *.xlsx; *.xlsm" '---엑셀 형식으로 지정
        .Show
        
        '파일 미 선택 시 종료 처리
        If .SelectedItems.count = 0 Then
        
            MsgBox "파일을 선택하지 않았습니다."
            Exit Sub
            
        '1개 파일 선택 시 기존 파일명 리스트 하위에 붙여넣기
        ElseIf .SelectedItems.count = 1 And 파일명.count < 10 And 파일명.count > 1 Then
            
            varFileNum = InStrRev(.SelectedItems(1), "\") '---'\' 기준으로 파일경로와 파일명 구분
            파일경로 = Left(.SelectedItems(1), varFileNum - 1) '---파일 경로 입력
            Cells(파일명.Row + 파일명.count, 2) = Mid(.SelectedItems(1), varFileNum + 1) '---파일명 입력
            Exit Sub
            
        End If
        
        파일경로.ClearContents '---파일 경로 초기화
        Range(파일명, 파일명.Offset(0, 1)).ClearContents '---파일명, 시트명 리스트 초기화
            
        Call SetRange '---파일명 영역 재지정
            
        For i = 1 To .SelectedItems.count
            
            '선택한 파일이 엑셀 형식이 아닌 경우 처리
            If InStr(.SelectedItems(i), ".xl") = 0 Then
            
                MsgBox "엑셀 파일을 선택해주세요."
                Exit Sub
                
            End If
            
            If i = 11 Then
                MsgBox "선택된 파일 개수가 10개를 초과하여 상위 10개의 파일 리스트만 호출됩니다."
                Exit For
            End If
            
            '파일명 리스트에 값 붙여넣기
            varFileNum = InStrRev(.SelectedItems(i), "\") '---'\' 기준으로 파일경로와 파일명 구분
            파일경로 = Left(.SelectedItems(i), varFileNum - 1) '---파일 경로 입력
            파일명(i) = Mid(.SelectedItems(i), varFileNum + 1) '---파일명 입력
            
        Next
    End With

End Sub
