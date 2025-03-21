Attribute VB_Name = "ModuleMain"
'=====================================================================
'매크로 : LoadFile
'대상 시트 : Home 시트
'동작 : 불러오기 버튼 동작, 입력되어 있는 파일 정보를 바탕으로 데이터 호출
'=====================================================================
Public Sub LoadFile()

    '###변수 선언###
    
    Dim var As Variant '---임시 변수
    Dim File_name_val As Variant '---확장자 제거한 파일명
    
    '###동작 시작###
    
    '---에러가 발생하면 종료
    On Error GoTo exit_error
    
    '---sub : 화면 업데이트 중지
    Call UpdateStart
    
    '---sub : 공통으로 사용하는 영역 위치 호출
    Call SetRange
    
    '---사용자 입력 영역 공백 시 알림 표시
    If File_adr = "" Or File_name = "" Or sheet_name = "" Then
    
        MsgBox "파일 정보를 모두 입력해주세요." & vbCrLf & "(파일 경로, 이름, 시트)"
        GoTo exit_sub

    End If
    
    '---function : 프리셋 공백 시 임시 이름 지정
    If preset = "" Or preset = "프리셋" Then
        
        preset = CheckPresetName
        프리셋명 = preset
    
    End If
    
    '---파일 이름에서 확장자(.xl~) 분리
    var = Split(File_name, ".")
    File_name_val = var(0)
    
    '---function : 프리셋 이름으로 시트, 쿼리 이미 생성되어 있다면 종료
    If CheckQuery = 1 Then
        
        MsgBox ("동일한 프리셋명이 존재합니다.")
        GoTo exit_sub
    
    '---function : 입력한 경로에 파일 존재 여부 체크
    ElseIf CheckFile(File_adr & "\" & File_name) = False Then
        
        MsgBox (File_adr & " 경로에 " & File_name & " 파일이 존재하지 않습니다.")
        GoTo exit_sub
    
    '---파일이 엑셀 형식이 아닌 경우 처리
    ElseIf InStr(File_name, ".xl") = 0 Then
     
         MsgBox "파일이 엑셀 형식이 아닙니다."
         GoTo exit_sub
         
    Else
    
        '---sub : 열 선택, 검색 영역 초기화
        Call ClearHomeData
                
        '---프리셋 이름으로 시트 생성
        ActiveWorkbook.Worksheets.Add after:=Sheets("Home")
        ActiveSheet.Name = preset
        
        '---입력된 경로를 바탕으로 파일 불러오기
        ActiveWorkbook.Queries.Add Name:=preset, _
        Formula:="let Source = Excel.Workbook(File.Contents(""" & File_adr & "\" & File_name & """), null, true), #""" & _
                sheet_name & "_Sheet"" = Source{[Item=""" & sheet_name & """, Kind=""Sheet""]}[Data], " & _
                "FilteredData = Table.PromoteHeaders(#""" & sheet_name & "_Sheet"") " & _
        "in FilteredData"

                
'                "모든열변경 = Table.TransformColumnTypes(FilteredData, " & _
'        "List.Transform(Table.ColumnNames(FilteredData), each {_, type text})) " & _
'        "in 모든열변경"
        
        '---연결된 쿼리 데이터 가져오기
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & preset & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & preset & "]")
            .Refresh BackgroundQuery:=False
        End With
        
        '---sub : 프리셋 저장
        Call preset_save
        
    End If
        
    Sheets("Home").Select
    
    '---파일 호출 시 열 선택 영역 강제 표시
    Sheets("Home").Columns("G:H").Hidden = False
    Sheets("Home").Shapes("Pic_Open").Visible = False
    Sheets("Home").Shapes("PIC_Close").Visible = True
    
    '---sub : 카테고리 리스트 호출
    Call SearchCategory(preset)
    
    '---function : 카테고리 전체 선택 및 추가
    If SelectAllCategory = 0 Then
        
        varCheckUpdate = Empty
        Call SelectAllCategory
        Call Button_AddCategory
        
    End If

'---종료 처리
exit_sub:
    
    Call UpdateEnd
    Range("A1").Select
    Exit Sub
    
'---에러 발생 처리
exit_error:
    
    '---시트명 오류 시 처리 (시트명 조건 체크, 쿼리명 조건 체크)
    If Left(Err.Description, 6) = "입력한 시트" Or Err.Number = -2147024809 Then
        
        '---시트 삭제 시 시스템 문구 미노출
        Application.DisplayAlerts = False
        
        '---생성된 시트 제거
        ActiveSheet.Delete
        Sheets("Home").Select
        
        Application.DisplayAlerts = True
        
        '---알림 표시
        MsgBox "프리셋명을 변경해주세요" & vbCrLf & vbCrLf & Err.Description
        GoTo exit_sub
        
    End If
    
    '---그 외 오류 처리
    MsgBox ("오류가 발생했습니다. 오류 코드 : " & Err.Number & vbCrLf & Err.Description)
    
    Call UpdateEnd
    Range("A1").Select
    
End Sub

'=====================================================================
'매크로 : SearchFile
'대상 시트 : Home 시트, etc 시트
'동작 : 파일 검색 버튼 클릭 시 동작, 파일 탐색기 오픈 후 사용자가 선택한 파일의 경로, 이름 호출
'=====================================================================
Public Sub SearchFile()
    
    '###변수 선언###
    
    Dim sel_File As Variant '---선택한 파일명 저장 변수
    Dim wb As Object '---가져온 파일 오브젝트 저장 변수
    Dim sheet_count As Variant '---불러온 파일의 시트 개수 저장 변수
    Dim file_open_check As String '---불러온 파일 오픈 여부 확인 변수
    Dim var() As String '---임시 사용 변수
    
    '###동작 시작###
    
    '--- sub 공통으로 사용하는 영역 위치, 색상 호출
    Call UpdateStart
    Call SetRange
    
    '---파일 주소 가져오기
    sel_File = File_adr & "\" & File_name
    
    '---etc시트 내 시트명 저장 영역 초기화
    Sheets("etc").Range("A:A").Clear
    
    '---파일 경로가 입력되어 있으면 해당 경로로 지정
    '---(잘못된 경로 입력 시 자동으로 무시됨)
    If File_adr <> "" Then
     
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = File_adr
         
    End If
    
    '---파일 탐색기 오픈
     With Application.FileDialog(msoFileDialogFilePicker)
         .Filters.Add "엑셀파일", "*.xls; *.xlsx; *.xlsm" '---엑셀 형식으로 지정
         .Show
         
         '---파일 미 선택 시 종료 처리
         If .SelectedItems.Count = 0 Then
         
             MsgBox "파일을 선택하지 않았습니다."
             GoTo exit_sub
             
         End If
         
         '---선택한 파일을 sel_File 변수에 저장
         sel_File = .SelectedItems(1)
         
     End With
     
     '---선택한 파일이 엑셀 형식이 아닌 경우 처리
     If InStr(sel_File, ".xl") = 0 Then
     
         MsgBox "엑셀 파일을 선택해주세요."
         GoTo exit_sub
         
     End If
                    
    '---파일 경로 및 이름 분리 후 저장
    max = InStrRev(sel_File, "\")
    파일경로.Value = Left(sel_File, max - 1)
    파일명.Value = Mid(sel_File, max + 1)
    
    '---sub : 파일 정보 재호출
    Call LoadFileInfo
           
    '---function : 사용자가 선택한 파일이 이미 열려있는지 확인
    file_open_check = CheckFileOpen(파일명.Value)
               
    '---선택한 파일의 시트명 불러오기
    Set wb = GetObject(sel_File)
    
    '---해당 파일의 시트 개수 확인
    sheet_count = wb.Sheets.Count
    
    '---시트 개수만큼 동작
    For n = 1 To sheet_count
    
        Sheets("etc").Range("a" & n) = wb.Sheets(n).Name
        
    Next
    
    '---sheet_list 이름 범위 재지정
    ThisWorkbook.Names("sheet_list").RefersTo = Sheets("etc").Range("A1", Sheets("etc").Cells(n - 1, 1))
    
    '---기존에 열려있지 않던 파일이라면 강제 종료
    If file_open_check = True Then
        
        wb.Close
        
    End If
    
    '---시트 목록 드롭 다운으로 표시
    With 시트명.Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=sheet_list"
    End With
    
    '---첫번째 시트를 기본값으로 표시
    시트명.Value = Sheets("etc").Range("A1").Value
    
'---종료 처리
exit_sub:

    Call UpdateEnd
    Range("A1").Select
    
End Sub

'=====================================================================
'매크로 : SearchCategory
'대상 시트 : Home 시트
'동작 : 파일 검색 버튼 클릭 시 동작, 파일 탐색기 오픈 후 사용자가 선택한 파일의 경로, 이름 호출
'=====================================================================
Public Sub SearchCategory(sheet_name)
    
    '###변수 선언###
        
    Dim category() As Variant   '---카데고리 값들을 저장할 배열
    Dim category_row As Range   '---불러온 데이터 시트에서 첫 행 영역 지정
    Dim Target As Range         '---현재 선택된 프리셋 명 저장 변수
    Dim category_range As Range '---카테고리 영역 지정
    
    '###동작 시작###
    
    현재프리셋.Value = sheet_name
    
    '---열 선택, 검색 영역 숨기기
    Call HideHomeData
    
    '---카테고리 리스트 영역 초기화
    열목록.Clear
    
    '---데이터 시트의 첫번째 행 영역 설정
    Set category_list = Sheets(sheet_name).Range("A1", Sheets(sheet_name).Range("A1").End(xlToRight))
    
    '---배열 크기 재정의
    ReDim category(category_list.Columns.Count - 1)
    
    '---배열에 첫번재 행 값 순차 저장
    For i = 0 To category_list.Columns.Count - 1
    
        category(i) = Sheets(sheet_name).Range("A1").Offset(0, i)
        
    Next
    
    '---Home 시트에 배열 값 뿌려주기
    For i = 0 To UBound(category)
    
        현재프리셋.Offset(i + 1, 0).Value = category(i)
    
    Next
    
    '---카테고리 영역 재 할당
    Call SetRange
    
    '---카테고리 영역 테두리 적용
    With 열목록
            
            .Borders.LineStyle = 1
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 1
        
    End With
    
    '---고정 행 영역 값 초기화
    고정행.ClearContents
        
End Sub

'=====================================================================
'매크로 : AddCategory
'대상 시트 : Home 시트
'동작 : 파일 검색 버튼 클릭 시 동작, 파일 탐색기 오픈 후 사용자가 선택한 파일의 경로, 이름 호출
'=====================================================================
Public Function AddCategory()

    'AddCategory = 1 : 에러 발생
    'AddCategory = 0 : 문제 없음
    
    '###변수 선언###
    
    Dim now_cell As Range
    Dim category_count As Variant
    Dim select_category As Range
    Dim search_row As Variant
    
    '###동작 시작###
    
    '---공통으로 사용하는 영역 위치, 색상 호출
    Call SetColor
    Call SetRange
    
    '---function : 카테고리가 존재하지 않으면 실행 종료
    If CheckCategory = 1 Then
    
        Range("notice") = "카테고리 리스트가 존재하지 않습니다."
        Range("notice").Font.Color = vbRed
        
        AddCategory = 1
        Exit Function
        
    End If
       
    Range("notice") = ""
    
    category_count = 0
    
    If 검색키워드_시작 <> Empty Then
        
        '---sub : 필터 여부 확인 후 초기화
        Call ResetSearch
            
        '---검색, 카테고리 영역 초기화
        Range("DATA").Clear
        
    End If
    
    '---선택된 카테고리 체크
    For i = 1 To 열목록.Rows.Count
        
        '---조회중인 셀 선언
        Set now_cell = 열목록_시작.Offset(i - 1, 0)
        
        '---색상으로 선택 여부 체크
        If Not now_cell.Interior.Color = vbWhite Then
            
            '---선택된 열 영역에 한 셀 씩 추가
            검색키워드_시작.Offset(0, category_count).Value = now_cell.Value
            
            '---선택된 열 영역 색 적용
            검색어_시작.Offset(0, category_count).Interior.Color = colorUserInput
            
            '---열 수량 체크
            category_count = category_count + 1
            
            '---select_category 변수에 선택된 열 주소 저장
            If select_category Is Nothing Then
            
                Set select_category = now_cell
                
            Else
            
                Set select_category = Union(select_category, now_cell)
                
            End If
        End If
    Next
    
    '---선택된 카테고리가 없는 경우
    If category_count = 0 Then
    
        Range("notice") = "선택된 카테고리가 없습니다."
        Range("notice").Font.Color = vbRed
        
        AddCategory = 1
        
    Else
    
        '---카테고리 추가된 열 너비 맞춤
        Range(검색키워드_시작, 검색키워드_시작.Offset(0, category_count - 1)).EntireColumn.AutoFit
        
        '---검색 영역 텍스트 타입으로 고정
        Range(검색어_시작, 검색어_시작.Offset(0, category_count - 1)).NumberFormatLocal = "@"
        
        AddCategory = 0
        
    End If
    
    With Sheets("etc")

        '---etc 시트내에서 프리셋 위치 찾기
        search_row = .Range("preset_list").Find(what:=현재프리셋.Value, lookat:=xlWhole).Row

        '---선택된 열 주소 프리셋에 저장하기
        If Not select_category Is Nothing Then

            .Cells(search_row, 6) = select_category.Address

        Else

            .Cells(search_row, 6).Clear

        End If

    End With

End Function

'=====================================================================
'매크로 : ResetCategory
'대상 시트 : Home 시트
'동작 : 열 선택 영역 선택 해제 버튼 동작, 열 리스트 전체 셀 서식 없애기
'=====================================================================
Public Sub ResetCategory()

    '---공통으로 사용하는 영역 위치 호출
    Call SetRange
    
    '---function : 표시된 열 존재하지 않으면 실행 종료
    If CheckCategory = 1 Then
    
        Range("notice") = "카테고리 리스트가 존재하지 않습니다."
        Range("notice").Font.Color = vbRed
        
        Exit Sub
    End If
    
    '---카테고리 영역 서식 초기화
    열목록.Interior.Color = vbWhite
    
End Sub

'=====================================================================
'매크로 : SelectAllCategory
'대상 시트 : Home 시트
'동작 : 열 선택 영역 전체 선택 버튼 동작, 열 리스트 전체 셀 서식 적용
'=====================================================================
Public Function SelectAllCategory()
    
    'SelectAllCategory = 1 : 표시된 열이 없음
    
    '---공통으로 사용하는 영역 위치, 색상 호출
    Call SetColor
    Call SetRange
    
    '---function : 표시된 열이 존재하지 않으면 실행 종료
    If CheckCategory = 1 Then
    
        Range("notice") = "카테고리 리스트가 존재하지 않습니다."
        Range("notice").Font.Color = vbRed
        SelectAllCategory = 1
        Exit Function
        
    End If
    
    '---열 전체 영역 셀 색상 적용
    열목록.Interior.Color = colorCategorySel
    SelectAllCategory = 0

End Function

'=====================================================================
'매크로 : ResetSearch
'대상 시트 : Home 시트
'동작 : 검색 영역 검색 초기화 버튼, 검색중이던 내용 초기화 및 데이터 시트 필터 해제
'=====================================================================
Public Function ResetSearch()
    
    'ResetSearch = 1 : 선택된 열이 없음
    'ResetSearch = 0 : 정상 동작
        
    '---공통으로 사용하는 영역 위치 호출
    Call SetRange
    
    '---검색란 초기화
    Range(검색어_시작, 검색키워드_시작.End(xlToRight).Offset(-1, 0)).ClearContents
    
    '---선택된 열 색 적용 초기화
    Range(검색키워드_시작, 검색키워드_시작.End(xlToRight)).ClearFormats
    
    '---선택된 열이 존재하지 않으면 실행 종료
    If 검색키워드_시작 = "" Then
    
        Range("notice") = "선택된 카테고리가 존재하지 않습니다."
        Range("notice").Font.Color = vbRed
        
        ResetSearch = 1
        
        Exit Function
        
    End If
    
    '---데이터 시트 영역에 필터가 걸려있다면 해제
    If Sheets(CStr(현재프리셋.Value)).AutoFilter.FilterMode = True Then

        Sheets(현재프리셋.Value).ShowAllData

    End If
    
    ResetSearch = 0
    
End Function
