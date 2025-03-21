Attribute VB_Name = "ModuleDesign"
'=====================================================================
'UI 디자인 관련 모듈
'=====================================================================

'###디자인 전역 변수###
Public colorCategorySel As Variant
Public colorUserInput As Variant

'=====================================================================
'좌측 메뉴 숨기기
Sub HideHomeMenu()

    Dim menu_range As Range '---메뉴 영역 선언
        
    Set menu_range = Sheets("Home").Columns("A:D")
    
    '---메뉴 현재 숨김 상태 확인 후 숨김, 표시 처리
    If menu_range.Hidden = True Then
    
        menu_range.Hidden = False
        
    Else
    
        menu_range.Hidden = True
                
    End If

End Sub

'=====================================================================
'색상 변수 선언
Public Sub SetColor()

    colorCategorySel = RGB(166, 201, 236) '---사용자가 선택 시 표시되는 색상
        
    colorUserInput = RGB(255, 178, 111) '---사용자가 입력할 수 있는 색상
    
End Sub

'=====================================================================
'검색된 데이터가 없을 때 데이터 영역 숨기기
Public Function HideHomeData()
       
    Dim data_range As Range '---데이터가 표시되는 영역 변수
    
    Call SetRange '---공통으로 사용하는 영역 위치 호출
    
    Set data_range = Sheets("Home").Columns("I:K") '---데이터 영역 저장
    
    '---현재 호출된 데이터 여부 체크 후 숨김, 표시 처리
    If 현재프리셋 = Empty Then
        
        data_range.Hidden = True
        
    Else
    
        data_range.Hidden = False
        
    End If
    
    Sheets("Home").Range("A1").Select
    
End Function

'=====================================================================
'열 선택 영역 숨기기
Public Sub HideHomeCategory()
    
    '---열 선택 영역 선언
    Dim rngCategory As Range
    
    Set rngCategory = Sheets("Home").Columns("G:H")
    
    Call SetRange '---공통으로 사용하는 영역 위치 호출
    
    If 현재프리셋 = Empty Then
        
        If rngCategory.Hidden = False Then
            
            Sheets("Home").Shapes("Pic_Open").Visible = True
            Sheets("Home").Shapes("PIC_Close").Visible = False
        
            rngCategory.Hidden = True
        End If
        
        Exit Sub
        
    End If
    
    '---메뉴 현재 숨김 상태 확인 후 숨김, 표시 처리
    If rngCategory.Hidden = True Then
    
        Sheets("Home").Shapes("Pic_Open").Visible = False
        Sheets("Home").Shapes("PIC_Close").Visible = True
        
        rngCategory.Hidden = False
        
    Else
    
        Sheets("Home").Shapes("Pic_Open").Visible = True
        Sheets("Home").Shapes("PIC_Close").Visible = False
        
        rngCategory.Hidden = True
                
    End If
    
End Sub
