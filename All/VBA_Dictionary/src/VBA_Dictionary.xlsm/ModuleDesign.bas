Attribute VB_Name = "ModuleDesign"
Option Explicit

Sub 이미지숨기기()
    
    Dim Shp1, Shp2 As Shape
    
    '이미지 지정
    Set Shp1 = Sheets("Main").Shapes("Cheat2_shp")
    Set Shp2 = Sheets("Main").Shapes("Cheat1_shp")
            
    Shp1.Visible = False
    Shp2.Visible = True
        
End Sub

Sub 셀테두리()

    '선택한 셀 테두리 적용
    With Target.Borders
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .Weight = xlMedium
    End With
    
    '테두리 적용 여부 확인
    If Target.Borders.LineStyle = xlContinuous Then
            
        '테두리 제거
        Target.Borders.LineStyle = xlNone
        
    End If
    
    
End Sub
