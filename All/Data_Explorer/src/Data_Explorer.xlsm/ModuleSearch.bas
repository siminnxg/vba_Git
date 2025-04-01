Attribute VB_Name = "ModuleSearch"
Option Explicit

Public Sub SaveSearch()
    
    Dim varSearchCol As Variant '---검색할 열 갯수
    Dim strSearchData As String '---입력된 검색어 저장
    Dim rngSearch As Range
    Dim rngKeyword As Range
    Dim varPresetIdx As Variant
    
    Call SetRange
    
    If 현재프리셋.Value = Empty Then
        Exit Sub
    End If
    
    Set rngKeyword = Range(검색키워드_시작, 검색키워드_끝)
    Set rngSearch = rngKeyword.Offset(-1, 0)
    
    varSearchCol = rngSearch.Cells.Count
    
    '검색되어 있는 내용 결합 후 변수에 저장
    strSearchData = rngKeyword(1) & "웷" & rngSearch(1) '---구분자로 웷 사용
    
    For i = 2 To varSearchCol
        
        strSearchData = strSearchData & "웷" & rngKeyword(i) & "웷" & rngSearch(i)
        
    Next
    
    If IsNull(strSearchData) = False Then
        With Sheets("etc")
            varPresetIdx = .Range("preset_list").Find(what:=현재프리셋.Value, lookat:=xlWhole).Row '---etc 시트내에서 프리셋 위치 찾기
    
            '검색중이던 데이터 프리셋에 저장하기
            .Cells(varPresetIdx, 7) = strSearchData
    
        End With
    End If
    
End Sub

Public Sub LoadSearch()
    
    Dim varSearchData As Variant
    Dim varPresetIdx As Variant
    Dim rngSel As Range
    
    On Error Resume Next
    
    Call SetRange
    
    With Sheets("etc")
    
        varPresetIdx = .Range("preset_list").Find(what:=현재프리셋.Value, lookat:=xlWhole).Row '---etc 시트내에서 프리셋 위치 찾기
        varSearchData = Split(.Cells(varPresetIdx, 7), "웷")
    
    End With
    
    '저장된 값 없으면 종료
    If UBound(varSearchData) = -1 Then
        Exit Sub
    End If
    
    Call UpdateStart
    Call SetColor
    
    For i = 0 To UBound(varSearchData)
        j = 0
        
        If varSearchData(i + 1) <> "" Then
            For Each rngSel In 검색키워드
            
                If rngSel.Value = varSearchData(i) Then
                    
                    rngSel.Offset(-1, 0).Value = varSearchData(i + 1)
                    rngSel.Interior.Color = colorCategorySel
                    
                    i = i + 1
                    Exit For
                    
                End If
                
                j = j + 1
            Next
        End If
    Next
    
    Call UpdateEnd
End Sub

Public Sub PasteData()
    
    Dim rngData As Range
    Dim varDataAry As Variant
    
    Call UpdateStart
    Call SetRange
    
    Set rngData = Sheets(현재프리셋.Value).ListObjects(1).Range
    
    Application.CutCopyMode = True
    
    With rngData.SpecialCells(xlCellTypeVisible)

        .Copy
        검색키워드_시작.PasteSpecial xlPasteValues

    End With

    Application.CutCopyMode = False
    
    'DATA 이름 범위 재설정
    ThisWorkbook.Names("DATA").RefersTo = Range(Selection, 검색어_시작)
    
    'DATA 이름 범위에 조건부 서식 적용
    With Range("DATA").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$F$5<>""""")
        
        '셀 테두리
        .Borders(xlLeft).LineStyle = xlContinuous
        .Borders(xlRight).LineStyle = xlContinuous
        .Borders(xlTop).LineStyle = xlContinuous
        .Borders(xlBottom).LineStyle = xlContinuous
        
    End With
    
    Sheets("Search").Range("A1").Select
    
    Call UpdateEnd
End Sub


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
    Range(검색어_시작, 검색키워드_끝.Offset(-1, 0)).ClearContents
    
    '---선택된 열 색 적용 초기화
    Range(검색키워드_시작, 검색키워드_끝).ClearFormats
    
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

Public Sub AutoFill(varIndex As Variant)
    
    Dim rngStart As Range
    Dim objData As Object
    Dim varSelCol As Variant
    
    Set objData = Sheets(현재프리셋.Value).ListObjects(1)
    varSelCol = objData.ListColumns(varIndex).Index
    
    Set rngStart = Sheets(현재프리셋.Value).Columns(varSelCol).Cells(1)
        
    Do Until rngStart.Row >= objData.Range.Rows.Count
        
        If rngStart.Offset(1, 0) = "" Then
            Range(rngStart, rngStart.End(xlDown).Offset(-1, 0)).FillDown
        
        End If
        
        Set rngStart = rngStart.End(xlDown)
        
    Loop
    
    Call PasteData

End Sub

