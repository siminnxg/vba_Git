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
    
    cnt = 0
    
    For Each cell In 검색목록.Offset(0, 9)
        If IsEmpty(cell) = False Then
            
            If 치트키_끝.Value = "" Then
                치트키_끝.Value = cell.Value
                
            Else
                치트키_끝.Offset(1, 0).Value = cell.Value
                
            End If
            
            
        End If
    Next
End Sub
