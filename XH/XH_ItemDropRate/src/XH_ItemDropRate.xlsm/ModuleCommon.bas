Attribute VB_Name = "ModuleCommon"
Option Explicit

Public 파일목록 As Range
Public 드랍율 As Range
Public 결과 As Range

Public i, j, k As Integer
Public cell As Range
Public control As control

'화면 업데이트 중지하여 실행 속도 증가
Public Sub UpdateStart()
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

'화면 업데이트 동작
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
    
     With Sheets("등급오류").UsedRange
        
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
                
    End With
    
    Sheets("Main").Range("A3").Select
    
End Sub

