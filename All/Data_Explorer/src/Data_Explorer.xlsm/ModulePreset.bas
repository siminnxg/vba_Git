Attribute VB_Name = "ModulePreset"
'=====================================================================
'매크로 : SavePreset
'대상 시트 : Home 시트, etc 시트
'동작 : 사용자가 입력한 파일 정보를 프리셋으로 저장
'=====================================================================
Public Sub SavePreset()
    
    '###변수 선언###
    
    Dim preset_count As Variant
    Dim file_info As Variant
        
    '###동작 시작###
        
    '---file_info 배열에 파일 정보 저장
    file_info = Array(preset, File_adr, File_name, sheet_name)
    
    '---빈칸 존재 시 알림창 표시
    If File_adr = "" Or File_name = "" Or sheet_name = "" Or preset = "" Then
    
        MsgBox "값을 모두 입력해주세요."
        Exit Sub
        
    End If
        
    With Sheets("etc")
        '---동일한 프리셋명 존재 시 알림창 표시
        If Not .Range("B:B").Find(what:=preset, lookat:=xlWhole) Is Nothing Then
        
            MsgBox "동일한 프리셋 명이 존재합니다."
            Exit Sub
            
        End If
        
        '---프리셋 개수 확인
        If .Range("B2") = "비어있음" Then
        
            preset_count = 1
            
        Else
        
            preset_count = .Range("B1").End(xlDown).Row
            
        End If
                    
        '---프리셋에 파일 정보 저장
        For i = 0 To 3
        
            .Cells(preset_count + 1, 2 + i).Value = file_info(i)
            
        Next
        
        '---preset_list 이름 범위 재지정
        ThisWorkbook.Names("preset_list").RefersTo = .Range("B1", .Cells(preset_count + 1, 2))
    
    End With
    
    '---슬라이서 새로고침
    ActiveWorkbook.SlicerCaches("슬라이서_프리셋").PivotTables(1).PivotCache.Refresh
    
End Sub
 
'=====================================================================
'매크로 : LoadPreset
'대상 시트 : Home 시트, etc 시트
'동작 : 선택된 프리셋 정보를 가져와서 데이터 호출
'=====================================================================
Public Sub LoadPreset()
    
    '###변수 선언###
    
    Dim varPresetSel As Variant '---현재 선택된 프리셋 값 저장 변수
    Dim varSearchRow As Variant '---etc 시트에서 현재 선택된 프리셋의 위치 저장 변수
    Dim varCategoryAdr As Variant '---프리셋에 저장되어있던 선택된 열들의 위치 저장 변수
        
    '###동작 시작###
    
    '화면 업데이트 중지
    Call UpdateStart
    
    '공통으로 사용하는 영역 위치, 색상 호출
    Call SetRange
    Call SetColor
    
    '에러가 발생하면 종료
    
    On Error GoTo exit_error
    
    With Sheets("etc").PivotTables("프리셋").DataBodyRange
        
        varPresetSel = CStr(.Cells(1)) '---현재 선택된 프리셋 값 변수에 저장
        
        '프리셋 관련 예외처리
        If varPresetSel = "Preset_Header" Then
            
            MsgBox ("프리셋을 선택해주세요,")
            GoTo exit_sub
                            
        ElseIf varPresetSel = "비어있음" Then
            
            MsgBox ("프리셋이 존재하지 않습니다.")
            GoTo exit_sub
        
        ElseIf .Cells.Count > 1 Then
            
            MsgBox ("프리셋을 1개만 선택해주세요.")
            GoTo exit_sub
            
        End If
    End With
    
    Call HideSearchSht(False)
    
    '기존 검색 데이터 저장 후 초기화
    Call SaveSearch
    Call ClearHomeData
    
    With Sheets("etc")
        'etc 시트내에서 프리셋 위치 찾기
        varSearchRow = .Range("preset_list").Find(what:=varPresetSel, lookat:=xlWhole).Row
        
        '프리셋 값 붙여넣기
        프리셋명 = .Cells(varSearchRow, 2).Value
        파일경로 = .Cells(varSearchRow, 3).Value
        파일명 = .Cells(varSearchRow, 4).Value
        시트명 = .Cells(varSearchRow, 5).Value
        varCategoryAdr = .Cells(varSearchRow, 6).Value
        
    End With
    
    '시트 목록 셀에 드롭다운 제거
    시트명.Validation.Delete
    
    Call SearchCategory(varPresetSel)
    
    '열 선택 상태를 저장해놓은 경우 호출
    If Not varCategoryAdr = "" Then
    
        Sheets("Search").Range(varCategoryAdr).Interior.Color = colorCategorySel
        
    End If
    
    '선택된 열 적용
    Call AddCategory
    
'---종료 처리
exit_sub:
    
    Call UpdateEnd
    Exit Sub
    
'---에러 발생 처리
exit_error:
    
    MsgBox ("오류가 발생했습니다. 오류 코드 : " & Err.Number & " " & Err.Description)
    Call UpdateEnd
End Sub

'=====================================================================
'매크로 : DeletePreset
'대상 시트 : Home 시트, etc 시트
'동작 : 선택된 프리셋을 제거
'=====================================================================
Public Sub DeletePreset()

    Dim varPresetSel As String
    Dim varSearchRow As Variant
    
    Call UpdateStart
    Call SetRange
    
    '---에러 발생해도 무시하고 진행
    On Error Resume Next
    
    '---시트 삭제 시 시스템 문구 미노출
    Application.DisplayAlerts = False
    
    With Sheets("etc").PivotTables("프리셋").DataBodyRange
        
        '---선택된 프리셋이 없는 때 처리
        If .Cells(1) = "Preset_Header" Then
        
            MsgBox ("프리셋을 선택해주세요.")
            GoTo exit_sub
        
        '---프리셋이 존재하지 않을 때 처리
        ElseIf .Cells(1) = "비어있음" Then
        
            MsgBox ("프리셋이 존재하지 않습니다.")
            GoTo exit_sub
        
        '---2개 이상 프리셋 선택 시 처리
        ElseIf .Cells.Count > 1 Then
        
            If MsgBox("프리셋이 두개 이상 선택되었습니다." & vbCrLf & "모두 제거하시겠습니까?", vbYesNo) = vbNo Then
            
                GoTo exit_sub
                
            End If
        End If
        
        '---선택된 프리셋 개수만큼 반복
        For i = 1 To .Cells.Count
            
            '--삭제 프리셋 목록에 현재 조회 중인 데이터가 있는 경우 조회중인 데이터 초기화
            If CStr(.Cells(i)) = 현재프리셋.Value Then
                
                '---검색 영역 초기화
                Call ClearHomeData
                
                '---열 선택 영역 초기화
                열목록.Clear
                현재프리셋.ClearContents
                
                Call HideSearchSht(True) '---검색 영역 숨기기
            End If
            
            '---프리셋 제거 시 시트, 쿼리 함께 제거
            Sheets(CStr(.Cells(i))).Visible = True
            Sheets(CStr(.Cells(i))).Delete
            ActiveWorkbook.Queries(CStr(.Cells(i))).Delete
            
            '---제거하려는 프리셋명 etc 시트에서 위치 검색
            varSearchRow = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc 프리셋 리스트 영역에서 프리셋 내용 제거
            Range(Range("preset_list")(varSearchRow), Range("preset_list")(varSearchRow).Offset(0, 5)).Delete Shift:=xlUp
            
            '---남아있는 프리셋 없는 경우 처리
            If Range("preset_list").Cells.Count = 1 Then
                
                Range("Preset_list").Offset(1, 0) = "비어있음"
                ThisWorkbook.Names("preset_list").RefersTo = Sheets("etc").Range("B1:B2") '---preset_list 이름 범위 재지정
                
                '---연결 전체 제거
                Call DeleteConnect
                
            End If
        Next
    End With
    
    '---슬라이서 프리셋 새로고침
    ActiveWorkbook.SlicerCaches("슬라이서_프리셋").PivotTables(1).PivotCache.Refresh
    
    '---데이터 이름 영역 재지정
    ThisWorkbook.Names("DATA").RefersTo = 검색어_시작
    
    '---시스템 문구 노출
    Application.DisplayAlerts = True

'종료 처리
exit_sub:
    Call UpdateEnd
    
End Sub


Public Sub EditPreset()
    
    Dim strFileAdr As String
    Dim varSearchRow As Variant
        
    Call UpdateStart
    
    With Sheets("etc").PivotTables("프리셋").DataBodyRange
        
        '---선택된 프리셋이 없는 때 처리
        If .Cells(1) = "Preset_Header" Then
        
            MsgBox ("프리셋을 선택해주세요.")
            GoTo exit_sub
        
        '---프리셋이 존재하지 않을 때 처리
        ElseIf .Cells(1) = "비어있음" Then
        
            MsgBox ("프리셋이 존재하지 않습니다.")
            GoTo exit_sub
            
        End If
        
        '---경로 입력 박스 표시
        strFileAdr = InputBox("변경할 파일 경로를 입력해주세요.", "프리셋 경로 변경")
        
        '---입력한 값이 없는 경우 종료
        If strFileAdr = Empty Then
            
            GoTo exit_sub
            
        End If
        '---선택된 프리셋 개수만큼 반복
        For i = 1 To .Cells.Count
            
            '---제거하려는 프리셋명 etc 시트에서 위치 검색
            varSearchRow = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc 프리셋 리스트 영역에서 프리셋 경로 변경
            Range("preset_list")(varSearchRow).Offset(0, 1).Value = strFileAdr
            
        Next
        
    End With
    
'종료 처리
exit_sub:
    Call UpdateEnd
    
End Sub


Public Sub RefreshPreset()
    
    Call SetRange
    
    With Range("preset_list")
        If .Cells.Count < 2 Then
            
            MsgBox "프리셋이 존재하지 않습니다."
            Exit Sub
            
        End If
        
        For i = 2 To .Cells.Count
            '입력한 경로에 파일 존재 체크
            If CheckFile(파일경로 & "\" & 파일명) = False Then
                
                MsgBox (파일경로 & " 경로에 " & 파일명 & " 파일이 존재하지 않습니다.")
            
            Else
                
                '프리셋 이름으로 생성된 시트의 listobject 새로고침
                Sheets(CStr(.Cells(i).Value)).ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
            
            End If
        Next
        
        MsgBox "최신 데이터를 갱신하였습니다."
        
    End With
End Sub
