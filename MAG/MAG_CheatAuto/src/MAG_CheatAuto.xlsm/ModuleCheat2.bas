Attribute VB_Name = "ModuleCheat2"
Option Explicit

'###################################################
'랜덤 옵션 아이템 생성 치트키 모듈
'###################################################


'======================================================================================================
'치트키2 [Cheat 생성] 버튼 클릭 시 동작


Public Sub Cheat2()
    
    Dim strCheatKey As String
    
    
    Call SetRange
    Call UpdateStart
           
    '선택된 KEY 가 없을 때 종료
    If IsEmpty(검색목록_시작) = True Then
            
        MsgBox "선택된 KEY가 존재하지 않습니다."
        
        Call UpdateEnd
        
        Exit Sub
            
    End If
    
    치트키.ClearContents '---치트키 영역 초기화
    
    '선택된 KEY 개수만큼 동작
    For Each cell In 검색목록.Offset(0, 10)
        
        Call SetRange
        
        '임시로 생성된 치트키가 없을 때 동작
        If IsEmpty(cell) = True Then
            
            If IsEmpty(cell.Offset(0, -8)) = False Then
            
                strCheatKey = "M1.Inven.RequestCreateEquipmentRandomOption " & cell.Offset(0, -8).Value & " 100 5 0 0 0 0 0 0 0 0"
            
            Else
                
                strCheatKey = "조회된 TID가 존재하지 않습니다."
                
            End If
            
        Else
        
            strCheatKey = cell.Value
            
        End If
        
        If 치트키_시작.Value = "" Then

            치트키_시작.Value = strCheatKey

        Else

            치트키_끝.Offset(1, 0).Value = strCheatKey

        End If
            
    Next
    
    '현재 프리셋 리스트 표시
    Call LoadTxt
    
    Call UpdateEnd
    
End Sub


Public Sub Cheat2TID()
    
    Dim strShtname As String
    Dim rngFind As Range
    
    Call SetRange
    
    For Each cell In 검색목록
    
        '아이템 타입 별 문서에서 KEY 검색 후 GroupId 추출
        For i = 1 To 3
        
            strShtname = 타입.ListColumns("문서").DataBodyRange(i).Value '---시트 순차적으로 지정
            
            Set rngFind = Sheets(strShtname).UsedRange.Find(cell.Value, Lookat:=xlWhole) '---선택된 셀을 검색할 시트에 검색
            
            '검색된 내용이 있을 때 동작
            If Not rngFind Is Nothing Then
            
                cell.Offset(0, 2) = rngFind.Offset(0, -1).Value '---TID 추출
                
                cell.Offset(0, 3) = rngFind.Offset(99, 1).Value '---100레벨 그룹 ID 추출
                
                Exit For '---검색 후 곧바로 반복 종료
                
            End If
        Next
    Next

End Sub
