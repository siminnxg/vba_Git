Attribute VB_Name = "ModuleMain"
Option Explicit

Public Sub SelectKey()
    
    Call SetRange
    
    '키 목록을 순회하며 선택된 셀 확인
    For Each cell In 키목록
        If cell.Borders.LineStyle = xlContinuous Then
            
            '검색 목록 비어있는 경우
            If 검색목록_시작.Value = "" Then
                검색목록_시작.Value = cell.Value
                
            '검색 목록 값이 존재하는 경우
            Else
                Set 검색목록_끝 = 검색목록_끝.Offset(1, 0)
                검색목록_끝.Value = cell.Value
            End If
        End If
    Next
    
    키목록.Borders.LineStyle = xlNone '---이동 완료 후 테두리 초기화
    검색어.Select

End Sub

Public Sub ClearKEYList1()

    Call SetRange
    
    검색목록.Resize(, 5).ClearContents
    
End Sub

Public Sub ClearKEYList2()

    Call SetRange
    
    검색목록.Borders.LineStyle = xlNone
    검색목록.Offset(0, 9).Resize(, 2).Clear
    검색목록.Resize(, 3).ClearContents
    
    Range(검색옵션_시작, 검색옵션_시작.End(xlDown)).Borders.LineStyle = xlNone
    Range(검색옵션_시작, 검색옵션_시작.End(xlDown)).ClearContents
        
    Range("Option").Offset(0, 1).Borders.LineStyle = xlNone
End Sub

Public Sub ClearCheatList()
    
    Call SetRange
    
    치트키.ClearContents
    프리셋.Offset(2, 0).Resize(100, 1).ClearContents
    
    치트키_시작.Offset(-1, 0).Value = "일괄 입력 희망 시 [메모장 생성] 버튼을 클릭해주세요."
    
End Sub

'Search 영역 초기화
Public Sub ClearSearchList()
    
    Call SetRange
    
    키목록.Borders.LineStyle = xlNone '---테두리 초기화
    
    검색어 = "" '---검색 내용 초기화
    
End Sub

Public Sub WriteCheat()
    
    Dim path As String
    Dim strPreset As String
    Dim strContents As String
    Dim lines() As String
    Dim check As Boolean
    Dim modifiedContent As String
    Dim index As Variant
    
    Call SetRange

    If IsEmpty(치트키_시작) Then
        MsgBox "생성된 치트키가 없습니다."
        Exit Sub
    End If
    
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    '프리셋명 확인
    If 프리셋 = "" Then
        strPreset = "<Mag_CreatItem>"
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
    
    If strPreset = "<Mag_CreatItem>" Then
        Open path For Input As #1
            strContents = Input$(LOF(1), 1)
        Close #1
                
        lines = Split(strContents, vbCrLf)
        
        If UBound(lines) = -1 Then
            GoTo 이어쓰기
        End If
        
        If lines(0) = "<Mag_CreatItem>" Then
            index = 1
        Else
            index = 0
        End If
            
        For i = index To UBound(lines) - 1
            If InStr(lines(i), "<") > 0 Then
                check = True
            End If

            If check = True Then
                    
                    modifiedContent = modifiedContent & lines(i)
                    
                    If i < UBound(lines) - 1 Then
                        
                        modifiedContent = modifiedContent & vbCrLf
                        
                    End If

            End If
        Next
        
        Open path For Output As #1
            Print #1, strPreset
            
            '치트키 영역을 돌면서 반복
            For Each cell In 치트키
                
                '조회된 TID~~ 셀 제외
                If InStr(cell.Value, "조회된") = 0 Then
                    Print #1, cell.Value '---작성된 치트 메모장에 입력
                End If
                
            Next
            
            Print #1, vbCrLf & vbCrLf & modifiedContent
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
        
종료:
    치트키_시작.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " """ & strPreset & """"
    
    Call LoadTxt
    
    Call UpdateEnd

End Sub

'Cheat 파일 열기
Public Sub OpenTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"

    If Dir(path, vbDirectory) = "" Then
        MsgBox "메모장을 생성해주세요."
        Exit Sub
    End If
    
    Shell "notepad.exe " & Chr(34) & path & Chr(34), vbNormalFocus

End Sub

'Cheat 파일에서 프리셋명을 찾아 리스트에 출력
Public Function LoadTxt()
    
    Dim path As String
    Dim strContents As String
    Dim lines() As String
    Dim strPresetList() As Variant
    
    Call SetRange
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
        
    프리셋.Offset(2, 0).Resize(1000, 1).ClearContents
    
    '생성되어있는 치트키 파일이 없는 경우 종료
    If Dir(path, vbDirectory) = "" Then
        Exit Function
    End If
    
    '파일에서 프리셋 명 읽기
    Open path For Binary As #1
        strContents = Space$(LOF(1))
        Get #1, , strContents
    Close #1
    
    lines = Split(strContents, vbCrLf)
    
    ReDim strPresetList(0 To 0)
    
    j = 0
    
    '프리셋명 리스트 추출
    For i = 0 To UBound(lines)
        If InStr(lines(i), "<") > 0 Then
            strPresetList(j) = lines(i)
            j = j + 1
            ReDim Preserve strPresetList(0 To j)
        End If
    Next
        
    For i = 0 To UBound(strPresetList)
        프리셋.Offset(2 + i, 0) = strPresetList(i)
    Next
    
    LoadTxt = 프리셋.Offset(2, 0).Resize(i, 1).Address
    
End Function
