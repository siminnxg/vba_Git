Attribute VB_Name = "ModuleDesign"
Option Explicit

Sub �̹��������()
    
    Dim Shp1, Shp2 As Shape
    
    '�̹��� ����
    Set Shp1 = Sheets("Main").Shapes("Cheat2_shp")
    Set Shp2 = Sheets("Main").Shapes("Cheat1_shp")
            
    Shp1.Visible = False
    Shp2.Visible = True
        
End Sub

Sub ���׵θ�()

    '������ �� �׵θ� ����
    With Target.Borders
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .Weight = xlMedium
    End With
    
    '�׵θ� ���� ���� Ȯ��
    If Target.Borders.LineStyle = xlContinuous Then
            
        '�׵θ� ����
        Target.Borders.LineStyle = xlNone
        
    End If
    
    
End Sub
