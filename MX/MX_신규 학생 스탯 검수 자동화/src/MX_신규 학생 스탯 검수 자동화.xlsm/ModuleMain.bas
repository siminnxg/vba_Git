Attribute VB_Name = "ModuleMain"
Option Explicit

Sub Student_Info()
    
    Dim 학생정보_시작 As Range
    Dim 학생요약_시작 As Range
    
    Dim cell1 As Range
    Dim cell2 As Range
    Dim rngFun As Range
    Dim cell_now As Range
    
    Dim cnt_header As Variant
    Dim cnt_st As Variant
    Dim stNum As Variant
    Dim column As String
    Dim i As Variant
    
    Dim ws As Worksheet
    
    Set ws = Sheets("데이터 처리")
    
    Set 학생정보_시작 = ws.Range("AF5")
    Set 학생요약_시작 = ws.Range("AB6")
    
    
    cnt_st = 0
    cnt_header = 0
    stNum = 1
    
    Set cell1 = 학생정보_시작
    
    Do While cell1 <> ""
        
        학생요약_시작.Offset(cnt_header, 0) = "학생" & stNum
                
        Set cell2 = cell1
        
        Do While cell2.Value <> "≪END≫"
            
            If InStr(cell2, "≪") Then
                
                '데이터 묶음 영역 지정
                Set rngFun = 학생요약_시작.Offset(cnt_header, 1)
                
                학생요약_시작.Offset(cnt_header, 1) = cell2
                
                '열 알파벳 추출
                column = Split(cell1.Offset(0, 3).Address(False, False), cell1.Offset(0, 3).Row)(0)
                
                '3열 함수 입력
                학생요약_시작.Offset(cnt_header, 2).Formula = "=LET(" & _
                                                                "titleCell, MATCH(" & rngFun.Address & ", " & column & ":" & column & ", 0)," & _
                                                                "nextTitleCell, IFERROR(AGGREGATE(15, 6, ROW($" & column & "$2:$" & column & "$100)/(LEFT($" & column & "$2:$" & column & "$100,1)=""≪"")/(ROW($" & column & "$2:$" & column & "$100)>titleCell), 1), ROWS(" & column & ":" & column & ")+1)," & _
                                                                "startRow, titleCell + 1," & _
                                                                "endRow, nextTitleCell - 1," & _
                                                                "colStart, COLUMN($" & column & "$2)," & _
                                                                "colEnd, colStart + 1," & _
                                                                "ADDRESS(startRow, colStart) & "":"" & ADDRESS(endRow, colEnd)" & _
                                                                ")"
                                
                
                '데이터 묶음 조회된 개수 증가
                cnt_header = cnt_header + 1
                
            End If
            
            Set cell2 = cell2.Offset(1, 0)
        Loop
                
        stNum = stNum + 1 '학생 수
        cnt_st = cnt_st + 5 '학생 정보 시작 위치
        
        Set cell1 = 학생정보_시작.Offset(0, cnt_st)
        
    Loop
    
End Sub

