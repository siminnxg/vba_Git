Attribute VB_Name = "ModuleCheat2"
Option Explicit

'###################################################
'랜덤 옵션 아이템 생성 치트키 모듈
'###################################################


'======================================================================================================
'치트키2 [Cheat 생성] 버튼 클릭 시 동작


Public Sub Cheat2()
    
    Call SetRange
    
    '선택된 KEY 개수만큼 동작
    For Each cell In 검색목록.Offset(0, 9)
        
        '임시로 생성된 치트키가 있을 때 동작
        If IsEmpty(cell) = False Then
            
            If 치트키_끝.Value = "" Then
            
                치트키_끝.Value = cell.Value
                
            Else
            
                치트키_끝.Offset(1, 0).Value = cell.Value
                
            End If
        End If
    Next
    
    '현재 프리셋 리스트 표시
    Call LoadTxt
    
End Sub
