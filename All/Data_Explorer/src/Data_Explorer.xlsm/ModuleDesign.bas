Attribute VB_Name = "ModuleDesign"
'=====================================================================
'UI 디자인 관련 모듈
'=====================================================================

'###디자인 전역 변수###
Public colorCategorySel As Variant
Public colorUserInput As Variant

'=====================================================================
'색상 변수 선언
Public Sub SetColor()

    colorCategorySel = RGB(166, 201, 236) '---사용자가 선택 시 표시되는 색상
        
    colorUserInput = RGB(255, 178, 111) '---사용자가 입력할 수 있는 색상
    
End Sub

'=====================================================================
'검색된 데이터가 없을 때 데이터 영역 숨기기
Public Function HideSearchSht(booCheck As Boolean)
    
    '시트 숨김, 표시 처리
    With Sheets("Search")
        If booCheck = True Then
            .Visible = 2
        Else
            .Visible = -1
            .Select
        End If
    End With
    
End Function

'=====================================================================
'열 선택 영역 숨기기
Public Sub HideHomeCategory(booCheck As Boolean)
        
    Dim rngCategory As Range '---열 선택 영역 선언
    
    Set rngCategory = Sheets("Search").Columns("B:C")
    
    Call SetRange
    
    '숨김, 표시 처리
    If booCheck = True Then
        rngCategory.Hidden = True
        
        '검색 영역 틀 고정
        검색키워드_시작.Offset(1, 틀고정.Value).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        Range("A1").Select
        
        
    Else
        rngCategory.Hidden = False
        
        '열 선택 영역 틀 고정
        Columns("D:D").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        Range("A1").Select
    End If
    
End Sub
