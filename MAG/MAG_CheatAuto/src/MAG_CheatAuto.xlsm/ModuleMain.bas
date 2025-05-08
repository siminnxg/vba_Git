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
    검색목록.Resize(, 3).ClearContents
    
    Range(검색옵션_시작, 검색옵션_시작.End(xlDown)).Borders.LineStyle = xlNone
    Range(검색옵션_시작, 검색옵션_시작.End(xlDown)).ClearContents
    
    Range("Option").Offset(0, 1).Borders.LineStyle = xlNone
End Sub

Public Sub ClearCheatList()
    
    Call SetRange
    
    치트키.ClearContents
    
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
    Dim fileContent As String
    Dim lines() As String
    Dim modifiedContent As String
    Dim check As Boolean
    
    Call SetRange
    
    If IsEmpty(치트키_시작) Then
        MsgBox "생성된 치트키가 없습니다."
        Exit Sub
    End If
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    'txt파일에 이어쓰기
    If 치트키_시작.Offset(0, 1) = False Then
    
        Open path For Append As #1
        
        Print #1, "<Mag_CreatItem>"
        
        '치트키 영역을 돌면서 반복
        For Each cell In 치트키
            
            '조회된 TID~~ 셀 제외
            If InStr(cell.Value, "조회된") = 0 Then
                Print #1, cell.Value '---작성된 치트 메모장에 입력
            End If
            
        Next
        
        Print #1, ""
        
        Close
    Else
        
        Open path For Input As #1
        fileContent = Input$(LOF(1), 1)
        Close #1
        
        lines = Split(fileContent, vbCrLf)
        
        For i = 0 To UBound(lines)
            If lines(i) = "" Then
                check = True
            End If
            
            If check = True Then
                
                modifiedContent = modifiedContent & lines(i) & vbCrLf
                
            End If
        Next
        
        Open path For Output As #1
        
        Print #1, "<Mag_CreatItem>"
        
        For Each cell In 치트키
        
            '조회된 TID~~ 셀 제외
            If InStr(cell.Value, "조회된") = 0 Then
                Print #1, cell.Value '---작성된 치트 메모장에 입력
            End If
            
        Next
        
        Print #1, modifiedContent
        Close #1
            
    End If
    
    치트키_시작.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " <Mag_CreatItem>"

End Sub
