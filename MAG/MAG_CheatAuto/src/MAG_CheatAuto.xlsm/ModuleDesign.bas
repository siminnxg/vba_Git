Attribute VB_Name = "ModuleDesign"
Option Explicit

Public Sub ChangeCheat()

    Dim Shp1 As Shape
    Dim Shp2 As Shape
    
    Set Shp1 = Sheets("Main").Shapes("Cheat1_shp")
    Set Shp2 = Sheets("Main").Shapes("Cheat2_shp")
    
    Call SetRange
    
    Range("A1").Select
    
    If rngCheat1.Hidden = True Then
        rngCheat1.Hidden = False
        rngCheat2.Hidden = True
        Shp1.Visible = False
        Shp2.Visible = True
    Else
        rngCheat1.Hidden = True
        rngCheat2.Hidden = False
        Shp1.Visible = True
        Shp2.Visible = False
    End If
    
End Sub



