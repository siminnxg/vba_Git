Attribute VB_Name = "ModuleCheat2"
Option Explicit

'###################################################
'랜덤 옵션 아이템 생성 치트키 모듈
'###################################################

'M1.Inven.RequestCreateEquipmentRandomOption (아이템TID) (레벨) (퍽 레벨) (옵션 1) (옵션 1 수치) (옵션 2) (옵션 2 수치) (옵션 3) (옵션 3 수치) (옵션 4) (옵션 4 수치)
Public Sub Cheat2()
    
    Dim strCheatKey As String
    Dim strCheatTid As String
    Dim strCheatStat As String
    
    Dim cnt As Variant
    
    Call SetRange
    
    If Range("option")(1) = "" Then
        Exit Sub
    End If
    
    i = 0
    
    '선택된 kEY 리스트 확인
    For Each cell In 검색목록
        If cell.Borders.LineStyle = xlContinuous Then
            strCheatKey = "M1.Inven.RequestCreateEquipmentRandomOption " & _
                            cell.Offset(0, 1).Value & " 100 5 "
            Exit For
        End If
    Next
    
    '선택된 옵션이 없는 경우 종료
    If IsNull(strCheatKey) Then
        Exit Sub
    End If
    
    For Each cell In Range("Option").Offset(0, 1)
    
        If cell.Borders.LineStyle = xlContinuous Then
            
            'TID 추출
            strCheatTid = cell.Offset(0, 1).Value
            
            'MAX 값 추출
            If 검색옵션_스텟 = False Then
                strCheatStat = cell.Offset(0, 4).Value
                
            'MIN 값 추출
            Else
                strCheatStat = cell.Offset(0, 3).Value
            End If
            
            strCheatKey = strCheatKey & strCheatTid & " " & strCheatStat & " "
            
            cnt = cnt + 1
            
        End If
        
    Next
    
    '옵션 최대 4개, 공백 시 0 0 입력
    For i = cnt To 3
        
        strCheatKey = strCheatKey & "0 0 "
        
    Next
    
    If 치트키_끝.Value = "" Then
        치트키_끝.Value = strCheatKey
    Else
        치트키_끝.Offset(1, 0).Value = strCheatKey
    End If
    
End Sub
