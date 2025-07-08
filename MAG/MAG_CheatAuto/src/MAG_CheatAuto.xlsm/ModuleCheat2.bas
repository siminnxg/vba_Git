Attribute VB_Name = "ModuleCheat2"
Option Explicit
'랜덤 옵션 아이템 생성 치트키 모듈


'######################################################################################################

'치트키2 [Cheat 생성] 버튼 클릭 시 동작

'######################################################################################################

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
            
                strCheatKey = "RequestCreateEquipmentAllOptions " & cell.Offset(0, -8).Value & " 100 4 True ()"
            
            Else
                
                strCheatKey = "조회된 TID가 존재하지 않습니다."
                
            End If
            
        Else
        
            strCheatKey = cell.Value
            
        End If
        
        '선택된 옵션, 코어 옵션 모두 없는 경우 처리
        If cell = "" And cell.Offset(0, 2) = "" Then
            
            strCheatKey = "M1.Inven.RequestCreateEquipmentRandomOption " & cell.Offset(0, -8).Value & " 100 5 0 0 0 0 0 0 0 0"
            
            GoTo 치트키입력
            
        End If
        
        '코어 옵션 체크되어 있을 때 동작
        If 코어체크 = True Then
            
            '선택된 코어 옵션이 없는 경우 true -> false 로 변경
            If cell.Offset(0, 2) = "" Then
                
                strCheatKey = Replace(strCheatKey, " True", " False")
                    
            '선택된 코어 옵션 있는 경우 치트키에 추가
            Else
            
                strCheatKey = strCheatKey & cell.Offset(0, 2)
            
            End If
        
        '코어 옵션 체크 해제 시 치트키 false 로 변경
        Else
            
            strCheatKey = Replace(strCheatKey, " True", " False")
            
        End If
        
치트키입력:
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

Public Sub Cheat2Core()
    
    Dim strCheatKey As String
    Dim strCoreAdr As String '---코어 리스트 입력 위치, 수량 저장 변수
    
    Application.EnableEvents = False
    
    strCheatKey = " ("
    
    '코어 옵션 리스트에서 수량 영역 순회
    For Each cell In Range("Core#").Resize(, 1).Offset(0, 4)
        
        '입력된 값이 있는 경우 동작
        If cell <> "" Then
                        
            Call CellBorder(cell.Offset(0, -4))
            
            '입력한 수량만큼 반복
            For i = 1 To cell.Value
                
                'max, min 상태에 따라 치트키 입력
                If Range("Core_Max") = True Then
                    
                    '코어 옵션 수치가 음수일 때 min <-> max 반대로 적용
                    If cell.Offset(0, -1) < 0 Then
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":min"","
                        
                    Else
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":max"","
                        
                    End If
                    
                Else
                    
                    '코어 옵션 수치가 음수일 때 min <-> max 반대로 적용
                    If cell.Offset(0, -1) < 0 Then
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":max"","
                        
                    Else
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":min"","
                    
                    End If
                End If
                
            Next
            
            strCoreAdr = strCoreAdr & cell.Address & "," & cell.Value & "&"
            
        Else
        
            cell.Offset(0, -4).Borders.LineStyle = xlNone
            
        End If
        
    
        
    Next
    
    strCheatKey = strCheatKey & ")"
    
    '입력한 수량이 없는 경우 동작
    If strCheatKey = " ()" Then
               
        strCheatKey = ""
        
        Cells(Range(Range("selKey").Value).Row, Range("Core").Column + 5).ClearContents
        
    End If
    
    '임시 치트키 입력
    Range(Range("selKey")).Offset(0, 11) = strCheatKey
    
    '입력된 위치와 수량 입력
    Cells(Range(Range("selKey").Value).Row, Range("Core").Column + 5) = strCoreAdr
    
End Sub

'======================================================================================================
'코어 리스트 분리

Public Sub CoreList()

    Dim rngCoreList As Range
    
    Set rngCoreList = Range("코어리스트").CurrentRegion.Columns(1)
        
    For Each cell In rngCoreList.Cells
        
        If InStr(1, cell.Value, "코어타입A") = 0 Then
            
            If InStr(1, cell.Offset(1, 0).Value, "코어타입A") > 0 Then
                
                ThisWorkbook.Names("코어_무기").RefersTo = Range(Range("코어리스트"), cell.Offset(0, 3))
                
                Exit For
                
            End If
            
        End If
    
    Next
        
    ThisWorkbook.Names("코어_외장부품").RefersTo = Range(cell.Offset(1, 3), cell.End(xlDown))
    
End Sub
