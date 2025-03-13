Attribute VB_Name = "ModulePreset"
'=====================================================================
'매크로 : preset_save
'대상 시트 : Home 시트, etc 시트
'동작 : 사용자가 입력한 파일 정보를 프리셋으로 저장
'=====================================================================
Public Sub preset_save()
    
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
'매크로 : preset_load
'대상 시트 : Home 시트, etc 시트
'동작 : 선택된 프리셋 정보를 가져와서 데이터 호출
'=====================================================================
Public Sub preset_load()
    
    '###변수 선언###
    
    Dim preset_select As Variant '---현재 선택된 프리셋 값 저장 변수
    Dim search_row As Variant '---etc 시트에서 현재 선택된 프리셋의 위치 저장 변수
    Dim category_adr As Variant '---프리셋에 저장되어있던 선택된 열들의 위치 저장 변수
        
    '###동작 시작###
    
    '---화면 업데이트 중지
    Call update_start
    
    '---공통으로 사용하는 영역 위치, 색상 호출
    Call range_set
    Call color_set
    
    '---에러가 발생하면 종료
    On Error GoTo exit_error
    
    With Sheets("etc").PivotTables("프리셋").DataBodyRange
    '---현재 선택된 프리셋 값 변수에 저장
        preset_select = CStr(.Cells(1))
        
        If preset_select = "Preset_Header" Then
            
            MsgBox ("프리셋을 선택해주세요,")
            GoTo exit_sub
                            
        ElseIf preset_select = "비어있음" Then
            
            MsgBox ("프리셋이 존재하지 않습니다.")
            GoTo exit_sub
        
        ElseIf .Cells.Count > 1 Then
            
            MsgBox ("프리셋을 1개만 선택해주세요.")
            GoTo exit_sub
            
        End If
    End With
    
    '--- sub : 검색, 열 선택 영역 초기화
    Call home_data_clear
    
    With Sheets("etc")
        '---etc 시트내에서 프리셋 위치 찾기
        search_row = .Range("preset_list").Find(what:=preset_select, lookat:=xlWhole).Row
        
        '---프리셋 값 붙여넣기
        user_file_preset = .Cells(search_row, 2).Value
        user_file_adr = .Cells(search_row, 3).Value
        user_file_name = .Cells(search_row, 4).Value
        user_file_sheet = .Cells(search_row, 5).Value
        category_adr = .Cells(search_row, 6).Value
        
    End With
    
    '---시트 목록 셀에 드롭다운 제거
    user_file_sheet.Validation.Delete
    
     '---입력한 경로에 파일 존재 체크
    If FileExists(user_file_adr & "\" & user_file_name) = False Then
        
        MsgBox (user_file_adr & " 경로에 " & user_file_name & " 파일이 존재하지 않아 기존 내용으로 불러옵니다.")
    
    Else
        
        '---프리셋 이름으로 생성된 시트의 listobject 새로고침
        Sheets(preset_select).ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    
    End If
    
    Call search_list(preset_select)
    
    '---열 선택 상태를 저장해놓은 경우 호출
    If Not category_adr = "" Then
    
        Sheets("Home").Range(category_adr).Interior.Color = category_sel_color
        
    End If
    
    '---선택된 열 적용
    Call button_category_add
    
'---종료 처리
exit_sub:
    
    Call update_end
    Exit Sub
    
'---에러 발생 처리
exit_error:
    
    MsgBox ("오류가 발생했습니다. 오류 코드 : " & Err.Number & " " & Err.Description)
    Call update_end
End Sub

'=====================================================================
'매크로 : preset_delete
'대상 시트 : Home 시트, etc 시트
'동작 : 선택된 프리셋을 제거
'=====================================================================
Public Sub preset_delete()

    Dim preset_select As String
    Dim search_row As Variant
    
    Call update_start
    Call range_set
    
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
            If CStr(.Cells(i)) = act_sheet_name.Value Then
                
                '---검색 영역 초기화
                Call home_data_clear
                
                '---열 선택 영역 초기화
                act_category_list.Clear
                act_sheet_name.ClearContents
                
                '---열 선택 영역 숨기기
                Call HideCategoryRng
                
            End If
            
            '---프리셋 제거 시 시트, 쿼리 함께 제거
            Sheets(CStr(.Cells(i))).Delete
            ActiveWorkbook.Queries(CStr(.Cells(i))).Delete
            
            '---제거하려는 프리셋명 etc 시트에서 위치 검색
            search_row = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc 프리셋 리스트 영역에서 프리셋 내용 제거
            Range(Range("preset_list")(search_row), Range("preset_list")(search_row).Offset(0, 5)).Delete Shift:=xlUp
            
            '---남아있는 프리셋 없는 경우 처리
            If Range("preset_list").Cells.Count = 1 Then
                
                Range("Preset_list").Offset(1, 0) = "비어있음"
                ThisWorkbook.Names("preset_list").RefersTo = Sheets("etc").Range("B1:B2") '---preset_list 이름 범위 재지정
                
                '---연결 전체 제거
                Call connect_delete
                
            End If
        Next
    End With
    
    '---슬라이서 프리셋 새로고침
    ActiveWorkbook.SlicerCaches("슬라이서_프리셋").PivotTables(1).PivotCache.Refresh
    
    '---데이터 이름 영역 재지정
    ThisWorkbook.Names("DATA").RefersTo = search_user_start
    
    '---검색 영역 숨기기
    Call home_data_hide
    
    '---시스템 문구 노출
    Application.DisplayAlerts = True

'종료 처리
exit_sub:
    Call update_end
    
End Sub


Public Sub preset_edit()
    
    Dim strFileAdr As String
    Dim search_row As Variant
        
    Call update_start
    
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
            search_row = Range("preset_list").Find(what:=.Cells(i), lookat:=xlWhole).Row
            
            '---etc 프리셋 리스트 영역에서 프리셋 경로 변경
            Range("preset_list")(search_row).Offset(0, 1).Value = strFileAdr
            
        Next
        
    End With
    
'종료 처리
exit_sub:
    Call update_end
    
End Sub

