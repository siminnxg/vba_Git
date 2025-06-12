Attribute VB_Name = "ModuleCommon"
Option Explicit

Public 파일목록 As Range
Public 드랍율 As Range
Public 결과 As Range

Public i, j, k As Integer
Public cell As Range
Public control As control


Public Sub UpdateStart()
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

Public Sub UpdateEnd()

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Public Sub RunForm()
    
    UserForm1.Show 0
    
End Sub

Public Sub ClearMain()

    With Sheets("Main").UsedRange
        
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
                
    End With
    
    Range("A3").Select
    
End Sub

