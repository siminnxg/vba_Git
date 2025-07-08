Attribute VB_Name = "ModuleDesign"
Option Explicit

'======================================================================================================
'치트키1 / 치트키2 변경


Public Sub ChangeCheat()

    Dim Shp1 As Shape
    Dim Shp2 As Shape
        
    '토글 버튼 이미지 지정
    Set Shp1 = Sheets("Main").Shapes("Cheat2_shp")
    Set Shp2 = Sheets("Main").Shapes("Cheat1_shp")
    
    Call SetRange
        
    Range("A1").Select
    
    If rngCheat1.Hidden = True Then
        rngCheat1.Hidden = False
        rngCheat2.Hidden = True
        Shp1.Visible = False
        Shp2.Visible = True
        코어체크 = False
        
    Else
        rngCheat1.Hidden = True
        rngCheat2.Hidden = False
        Shp1.Visible = True
        Shp2.Visible = False
    End If
    
End Sub

Public Sub CellBorder(rngBorder As Range)
    
    With rngBorder.Borders
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .Weight = xlMedium
    End With
    
End Sub

