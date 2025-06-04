Attribute VB_Name = "ModuleMain"
Option Explicit

'======================================================================================================
'Search 영역에서 [KEY 선택] 버튼 클릭 시 동작


Public Sub SelectKey()
    
    Call SetRange
    
    i = 2
    j = ""
    
    If 검색어.Offset(0, 1) = "커스터마이징" Then
        
        i = 3
        j = "CustomizingItemData"
    
    End If
    
    '키 목록을 순회하며 선택된 셀 확인
    For Each cell In 키목록
    
        '현재 셀에 테두리가 적용된 상태일 때 동작
        If cell.Borders.LineStyle = xlContinuous Then
            
            '검색 목록 첫 행이 비어있을 때 동작
            If 검색목록_시작.Value = "" Then
            
                검색목록_시작.Resize(1, i) = cell.Resize(1, i).Value '---첫 행에 KEY 입력
                
                검색목록_시작.Offset(0, i) = j
                
            '검색 목록 첫 행에 값이 존재할 때 동작
            Else
            
                Set 검색목록_끝 = 검색목록_끝.Offset(1, 0)
                
                검색목록_끝.Resize(1, i) = cell.Resize(1, i).Value '---끝 행에 KEY 입력
                
                검색목록_끝.Offset(0, i) = j
                
            End If
        End If
    Next
    
    If rngCheat2.Hidden = False And IsEmpty(검색목록_시작) = False Then
        
        Call Cheat2TID
        
    End If
    
    키목록.Borders.LineStyle = xlNone '---이동 완료 후 테두리 초기화
    
    검색어.Select '---검색어 영역에 커서 이동
       
End Sub

'======================================================================================================
'치트키1 [초기화] 버튼 클릭 시 동작

 
Public Sub ClearKEYList1()

    Call SetRange
    
    검색목록.Resize(, 6).ClearContents
    
End Sub


'======================================================================================================
'치트키2 [초기화] 버튼 클릭 시 동작


Public Sub ClearKEYList2()

    Call SetRange
    
    검색목록.Offset(0, 1).Borders.LineStyle = xlNone
    검색목록.Offset(0, 10).Resize(, 2).Clear
    검색목록.Resize(, 4).ClearContents
    
    Range(검색옵션_시작, 검색옵션_시작.End(xlDown)).Borders.LineStyle = xlNone
    Range(검색옵션_시작, 검색옵션_시작.End(xlDown)).ClearContents
        
    Range("Option").Offset(0, 1).Borders.LineStyle = xlNone
End Sub

'======================================================================================================
'치트키 리스트 [초기화] 버튼 클릭 시 동작


Public Sub ClearCheatList()
    
    Call SetRange
    
    치트키.ClearContents
    
    치트키_시작.Offset(-1, 0).Value = "일괄 입력 희망 시 [메모장 입력] 버튼을 클릭해주세요."
    
End Sub

'======================================================================================================
'Search 영역 [초기화] 버튼 클릭 시 동작


Public Sub ClearSearchList()
    
    Call SetRange
    
    키목록.Borders.LineStyle = xlNone '---테두리 초기화
    
    검색어 = "" '---검색 내용 초기화
    
End Sub

'======================================================================================================
'[메모장 입력] 버튼 클릭 시 동작


Public Sub WriteCheat()
    
    '# 변수 선언 #
    
    Dim path As String '---파일 경로 저장 변수
    Dim strPreset As String '---프리셋명 저장 변수
    Dim strContents As String '---메모장 내용 저장 변수
    Dim lines() As String '---메모장 내용을 줄바꿈 단위로 구분하여 저장 변수
    Dim check As Boolean '---입력된 프리셋명 체크 변수
    Dim modifiedContent As String '---기존 입력되어 있던 내용 백업 변수
    
    
    '# 동작 시작 #
    
    Call SetRange
    
    '치트키 목록이 비어있으면 종료
    If IsEmpty(치트키_시작) Then
    
        MsgBox "생성된 치트키가 없습니다."
        
        Exit Sub
        
    End If
    
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt" '---파일 경로 지정
    
    '프리셋명 공백일 때 고정값 입력
    If 프리셋 = "" Then
    
        strPreset = "<Mag_CreateItem>"
        
    Else
    
        strPreset = "<" & 프리셋.Value & ">"
        
    End If
        
    '생성된 파일이 없는 경우 신규 생성
    If Dir(path, vbDirectory) = "" Then
        
        Open path For Output As #1
                        
            Print #1, strPreset
            
            '치트키 영역을 돌면서 반복
            For Each cell In 치트키
                
                '조회된 TID~~ 셀 제외
                If InStr(cell.Value, "조회된") = 0 Then
                
                    Print #1, cell.Value '---작성된 치트 메모장에 입력
                    
                End If
                
            Next
            
            Print #1, vbCrLf
        Close
        
        GoTo 종료
    
    End If
    
    '프리셋 명이 공백일 때 동작
    '<Mag_CreateItem> 치트키는 입력된 내용으로 덮어쓰고 하위에 다른 프리셋명은 백업 후 붙여넣기
    If strPreset = "<Mag_CreateItem>" Then
    
        '메모장 읽기
        Open path For Input As #1
        
            strContents = Input$(LOF(1), 1)
            
        Close #1
                
        lines = Split(strContents, vbCrLf) '---줄바꿈을 단위로 구분하여 배열에 저장
        
        '입력된 값이 없는 경우 값 입력으로 이동
        If UBound(lines) = -1 Then
        
            GoTo 이어쓰기
            
        End If
        
        'i = MsgBox("<Mag_CreateItem> 프리셋을 덮어쓰시겠습니까?", vbYesNo) '---덮어쓰기 여부 문의
        
        '사용자가 NO 선택 시 종료
        If MsgBox("<Mag_CreateItem> 프리셋을 덮어쓰시겠습니까?", vbYesNo) = 7 Then
        
            GoTo 종료
            
        End If
        
        '<Mag_CreateItem> 를 제외하고 프리셋명이 입력된 행 백업
        For i = 0 To UBound(lines) - 1
            
            '프리셋 명 확인
            If InStr(lines(i), "<") > 0 And lines(i) <> "<Mag_CreateItem>" Then
            
                check = True
                
            End If
            
            If check = True Then
                    
                    modifiedContent = modifiedContent & lines(i) '---해당 변수에 백업
                    
                    '마지막 행일 때 동작
                    If i < UBound(lines) - 1 Then
                        
                        modifiedContent = modifiedContent & vbCrLf '---줄바꿈 추가
                        
                    End If
            End If
        Next
        
        '메모장 새로 쓰기
        Open path For Output As #1
        
            Print #1, strPreset
            
            '치트키 영역을 돌면서 반복
            For Each cell In 치트키
                
                '조회된 TID~~ 셀 제외
                If InStr(cell.Value, "조회된") = 0 Then
                
                    Print #1, cell.Value '---작성된 치트 메모장에 입력
                    
                End If
                
            Next
            
            Print #1, vbCrLf & vbCrLf & modifiedContent '---하위에 백업해둔 내용 입력
            
        Close
        
        GoTo 종료
        
    End If
    
    '동일한 프리셋 명 체크
    For Each cell In Range(LoadTxt)
        
        If cell.Value = strPreset Then
        
            MsgBox strPreset & " : 동일한 프리셋 명이 존재합니다."
            
            Call UpdateEnd
            
            Exit Sub
            
        End If
        
    Next
    
이어쓰기:

    'txt파일에 이어쓰기
    Open path For Append As #1
    
        Print #1, strPreset '---프리셋명 입력
            
        '치트키 영역을 돌면서 반복
        For Each cell In 치트키
            
            '조회된 TID~~ 셀 제외
            If InStr(cell.Value, "조회된") = 0 Then
            
                Print #1, cell.Value '---작성된 치트 메모장에 입력
                
            End If
        Next
        
        Print #1, vbCrLf
        
    Close
        
종료:
    치트키_시작.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " """ & strPreset & """"
    
    Call LoadTxt
    
    Call UpdateEnd

End Sub

'======================================================================================================
'[메모장 열기] 버튼 클릭 시 동작


Public Sub OpenTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    '경로에 메모장 파일이 없으면 종료
    If Dir(path, vbDirectory) = "" Then
    
        MsgBox "메모장을 생성해주세요."
        
        Exit Sub
        
    End If
    
    Shell "notepad.exe " & Chr(34) & path & Chr(34), vbNormalFocus '---메모장 열기

End Sub

'======================================================================================================
'Cheat 파일에서 프리셋명을 찾아 리스트에 출력


Public Function LoadTxt()
    
    Dim path As String
    Dim strContents As String
    Dim lines() As String
    Dim strPresetList() As Variant
    
    Call SetRange
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
        
    프리셋.Offset(2, 0).Resize(1000, 1).ClearContents
    
    '생성되어있는 치트키 파일이 없는 경우 종료
    If Dir(path, vbDirectory) = "" Then
    
        Exit Function
        Call UpdateEnd
        
    End If
    
    '파일에서 프리셋 명 읽기
    Open path For Binary As #1
    
        strContents = Space$(LOF(1))
        
        Get #1, , strContents
        
    Close #1
    
    lines = Split(strContents, vbCrLf) '---줄바꿈을 기준으로 분리
    
    ReDim strPresetList(0 To 0)
    
    j = 0
    
    '프리셋명 리스트 추출
    For i = 0 To UBound(lines)
        
        '프리셋 명 확인 시 동작
        If InStr(lines(i), "<") > 0 Then
        
            strPresetList(j) = lines(i)
            
            j = j + 1
            
            ReDim Preserve strPresetList(0 To j) '---행 수 만큼 배열 확장
            
        End If
    Next
        
    For i = 0 To UBound(strPresetList)
    
        프리셋.Offset(2 + i, 0) = strPresetList(i)
        
    Next
    
    LoadTxt = 프리셋.Offset(2, 0).Resize(i, 1).Address
    
    Set 프리셋_끝 = 프리셋.Offset(1 + i, 0)
    
    Call UpdateEnd
    
End Function

'======================================================================================================
'[메모장 초기화] 버튼 클릭 시 동작


Public Sub ClearTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt" '---메모장 경로 지정
    
    '생성된 파일이 없을 때 종료
    If Dir(path, vbDirectory) = "" Then
    
        MsgBox "생성된 파일이 존재하지 않습니다."
        
        Exit Sub
        
    End If
    
    If MsgBox("메모장을 초기화 하시겠습니까?", vbYesNo) = vbYes Then
    
        'txt파일 초기화
        Open path For Output As #1
        Close
        
        '프리셋 리스트 갱신
        Call LoadTxt
        
        치트키_시작.Offset(-1, 0).Value = "일괄 입력 희망 시 [메모장 입력] 버튼을 클릭해주세요."
    
    End If
    
End Sub
