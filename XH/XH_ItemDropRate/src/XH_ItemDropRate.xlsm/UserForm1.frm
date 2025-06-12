VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ItemDropRate"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8385.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'======================================================================================================
'드랍율 세팅 [기본 값] 버튼 클릭 시 동작

Private Sub Button_BaseRate_Click()
    
    Me.드랍율1.Value = "0.81"
    Me.드랍율2.Value = "0.11"
    Me.드랍율3.Value = "0.0081"
    Me.드랍율4.Value = "0.00081"
    Me.드랍율5.Value = "0.0000331"
    Me.드랍율6.Value = "1"
    
End Sub

'======================================================================================================
'문서 리스트 [문서 선택] 버튼 클릭 시 동작

Private Sub Button_FileLoad_Click()
    
    Dim strFileName As String '---선택한 파일의 경로 및 이름 저장 변수
    
    '파일 탐색기 오픈
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "엑셀파일", "*.xls; *.xlsx; *.xlsm" '---엑셀 형식으로 지정
        .Show
        
        '파일 미 선택 시 종료 처리
        If .SelectedItems.Count = 0 Then
        
            MsgBox "파일을 선택하지 않았습니다."
            
            Exit Sub
            
        End If
        
        'etc 시트 '파일이름' 영역 초기화
        Range("파일이름").ClearContents
        ThisWorkbook.Names("파일이름").RefersTo = Sheets("etc").Range("B3")
                        
        '파일 리스트에 선택된 파일 경로 입력
        For i = 1 To .SelectedItems.Count

            Me.List_File.AddItem .SelectedItems(i)

        Next
        
        'etc 시트 '파일이름' 영역에 경로를 제외한 파일 이름만 입력
        For i = 0 To Me.List_File.ListCount - 1
            
            strFileName = InStrRev(Me.List_File.List(i), "\") '---파일 경로와 분리
            
            Range("파일이름").Offset(i, 0).Value = Mid(Me.List_File.List(i), strFileName + 1) '---파일 이름 입력
            
        Next
        
        '파일이름' 영역 재지정
        ThisWorkbook.Names("파일이름").RefersTo = Range("파일이름").Resize(i + 1, 1)
              
    End With
    
End Sub

'======================================================================================================
'문서 리스트 [초기화] 버튼 클릭 시 동작

Private Sub CommandButton2_Click()
    
    Me.List_File.Clear
    
End Sub

'======================================================================================================
'[조회] 버튼 클릭 시 동작

Private Sub CommandButton1_Click()
    
    Dim cntResult As Variant
    
    Dim objDB As Object
    Dim obj As Object
    Dim strSql As String
    
    Dim strLevel As String
    Dim varDropRate As Variant
    Dim path As String
        
    Dim 결과 As Range '---드랍율을 초과한 데이터를 출력할 위치
    Dim 결과2 As Range '---아이템 등급이 일치하지 않는 데이트를 출력할 위치
    Dim 파일명 As Range
    
    '#동작 시작#
    
    '에러 발생 시 무시
    On Error Resume Next
    
    k = 1
    
    '문서 선택 여부 확인
    If IsNull(Me.List_File.List(0)) Then
    
        MsgBox "선택된 문서가 없습니다."
        Exit Sub
        
    End If
    
    '드랍율 공백 및 숫자 형식 확인
    For i = 1 To 6
        
        Set control = Me.Controls("드랍율" & i)
        
        If control = "" Then
            
            MsgBox "드랍율을 입력해주세요."
            Exit Sub
            
        ElseIf IsNumeric(control) = False Then
            
            MsgBox "드랍율을 숫자 형식으로 입력해주세요."
            Exit Sub
            
        End If
        
    Next
    
    '드랍율 오차 숫자 형식 확인
    If IsNumeric(Me.드랍율오차) = False Then
        
        MsgBox "드랍율 오차 수치를 숫자 형식으로 입력해주세요."
        Exit Sub
            
    End If
    
    '화면 업데이트 중지
    Call UpdateStart
    
    'Main 시트 초기화
    Call ClearMain
    
    Set 결과 = Sheets("Main").Range("B4") '---붙여 넣을 영역 지정
        
    Set 결과2 = Sheets("등급오류").Range("B2")
    
    Range("헤더").Copy Destination:=결과2.Offset(-1, 0) '--머릿글 행 추가
    
    For i = 0 To Me.List_File.ListCount - 1
        
        cntResult = 0
                
        path = Me.List_File.List(i) '---문서 리스트 순차적으로 경로 지정
        
        Set 파일명 = 결과.Offset(-1, -1)
        
        파일명 = Range("파일이름")(i + 1) '---조회된 파일명 입력
        
        결과2.Offset(0, -1).Value = Range("파일이름")(i + 1)
        
        'OLEDB 연결
        Set objDB = CreateObject("ADODB.Connection")
        objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=" & path & ";" & _
                  "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
                
        '등급 수 만큼 반복 (6개)
        For j = 1 To Range("등급").Cells.Count
            
            strLevel = Range("등급")(j).Value '---등급 변수에 저장
            
            Set control = Me.Controls("드랍율" & j) '---드랍율 텍스트 박스 순차 지정
            
            varDropRate = control.Value * 0.01 '---드랍율 변수에 저장
            
            '(아이템 타입, 등급 일치, 드랍율을 초과) 조건을 만족하는 데이터 조회
            strSql = " SELECT * FROM [Data$] WHERE F10 = '아이템' AND F8 LIKE '%" & strLevel & "%' AND F9 = '" & strLevel & "' AND F12 LIKE '%Grade" & j & "%' AND F17 > " & varDropRate
            
            Set obj = objDB.Execute(strSql)
            
            '조회된 데이터가 없으면 다음 등급으로 이동
            If obj.EOF Then
                
                GoTo 등급조회
            
            '조회된 데이터가 있는 경우 Main 시트에 입력
            Else
                
                결과.CopyFromRecordset obj '---조회된 데이터 입력
                
                Range("헤더").Copy Destination:=결과.Offset(-1, 0) '---머릿글 입력
                
                '드랍율 오차 확인
                Call CheckRate(결과.CurrentRegion, varDropRate)
                
                Set 결과 = 결과.End(xlDown).Offset(3, 0) '---결과 영역 재지정
                
                cntResult = 1 '---조회된 데이터 여부 확인
                
                Application.CutCopyMode = False
                
            End If
            
등급조회:
            '아이템 타입 중 등급이 일치하지 않는 조건 조회
            strSql = " SELECT * FROM [Data$] WHERE F10 = '아이템' AND F8 LIKE '%" & strLevel & "%' AND (F9 <> '" & strLevel & "' OR F12 NOT LIKE '%Grade" & j & "%')"
            
            Set obj = objDB.Execute(strSql)
            
            '조회된 데이터가 없으면 다음 등급으로 이동
            If obj.EOF Then
                
                GoTo 다음등급
            
            '조회된 데이터가 있는 경우 등급오류 시트에 입력
            Else
                
                결과2.CopyFromRecordset obj '---조회된 데이터 입력
                
                Set 결과2 = 결과2.End(xlDown).Offset(1, 0)
                
                Application.CutCopyMode = False
                
            End If
                        
다음등급:

        Next

        '개체 연결 끊기
        obj.Close
        objDB.Close
        Set obj = Nothing
        Set objDB = Nothing
        
        '조회된 데이터가 없으면 파일명 제거
        If cntResult = 0 Then
        
            파일명.ClearContents
            
        End If
        
        j = j + 1
        
    Next
        
    Columns("R:R").NumberFormatLocal = "0.000000%" '---드랍율 서식 변경
    
    Sheets("Main").UsedRange.Columns.AutoFit '---결과 자동 열 맞춤
    
    '조회 종료 시 알림
    If Range("B5") = "" Then
        
        MsgBox "조회된 데이터가 없습니다."
        
    Else
        
        MsgBox "조회 완료되었습니다."
        
    End If

종료:
    Call UpdateEnd
    
End Sub

'======================================================================================================
'드랍율 오차 확인

Public Function CheckRate(rngResult As Range, varDropRate As Variant)
    
    Dim varRate As Variant '---드랍율 오차 수치 저장 변수
    
    '현재 조회된 데이터 테두리 적용
    With rngResult.Borders
        
        .LineStyle = xlContinuous
        .Color = rgbGainsboro
        
    End With
    
    '수치 입력되어 있는 경우 동작
    If Me.드랍율오차.Value <> "" Then
        
        varRate = Me.드랍율오차.Value * 0.01 '---드랍율 오차 수치 변수에 입력
        
        '조회 결과 영역을 순회
        For Each cell In rngResult
                        
            '드랍율 표시 열, 드랍율 오차 수치 + 드랍율 수치 보다 큰 경우 행 색상 적용
            If cell.Column = 18 And cell.Value > (varRate + varDropRate) Then
            
                Range(Cells(cell.Row, 2), Cells(cell.Row, cell.Column)).Interior.ColorIndex = 6
                
            End If
                    
        Next
        
    End If

End Function
