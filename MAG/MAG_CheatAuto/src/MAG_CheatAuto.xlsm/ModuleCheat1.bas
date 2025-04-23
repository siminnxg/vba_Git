Attribute VB_Name = "ModuleCheat1"
Option Explicit
'###################################################
'아이템 생성 치트키 모듈
'###################################################

Sub FindTID()
    
    '# 변수 선언
    Dim rngFileName As Range '파일경로 & 파일명
    Dim rngFindType As Range
    
    '# 동작 시작
    Call UpdateStart
    Call SetRange
    
    If 검색목록_시작.Value = "" Then
        MsgBox "선택된 Key가 없습니다."
        GoTo Exit_Sub
    End If
        
    Call SQLFileLoad(검색목록, 타입.ListColumns("문서").DataBodyRange)
    
    Call CheatCreatItem
    
    치트키_시작.Offset(-1, 0).Value = "일괄 입력 희망 시 [메모장 생성] 버튼을 클릭해주세요."
    
Exit_Sub:
    Call UpdateEnd
    
End Sub

Public Sub SelectKey()
    
    Call SetRange
    
    '키 목록을 순회하며 선택된 셀 확인
    For Each cell In 키목록
        If cell.Offset(0, -1).Interior.Color = vbRed Then
            
            '검색 목록 비어있는 경우
            If 검색목록_시작.Value = "" Then
                검색목록_시작.Value = cell.Value
                
            '검색 목록 값이 존재하는 경우
            Else
                Set 검색목록_끝 = 검색목록_끝.Offset(1, 0)
                검색목록_끝.Value = cell.Value
            End If
        End If
    Next
    
    키목록.Offset(0, -1).Interior.Color = vbWhite '---이동 완료 후 KEY 색상 초기화

End Sub

Public Function SQLFileLoad(cell As Range, rngFileName As Range)
    
    '# 변수 선언
    Dim objDB As Object 'ADODB 개체 생성할 변수
    Dim obj As Object '데이터 개체 담을 변수
    Dim strSQL As String 'SQL문 담을 변수
    Dim strFilePath As String '파일 경로
    
    Dim strWhere As String '---Where 조건 변수
    Dim rngFindCell As Range '---검색된 Key의 위치 변수
    Dim rngRuneCell As Range '---RuneData 시트에서 key의 위치 변수
    Dim strRuneData As Variant '---RuneUIData 문서에서 찾은 값 저장 배열
    Dim strFileName As String
    
    '# 동작 시작
    'On Error Resume Next
    
    '선택된 Key 값들을 묶어 Where 조건으로 변환
    For i = 1 To cell.Cells.Count
        strWhere = strWhere & "'" & cell(i).Value & "',"
    Next
    
    '문서 개수만큼 반복
    For i = 1 To rngFileName.Cells.Count
        
        strFileName = rngFileName(i).Value
        
        strFilePath = 파일경로.Value & "\" & strFileName & ".xlsx" '---문서 경로 지정
        
        'SQL 조건 작성
        '룬 데이터는 개별 작성
        If rngFileName(i) = "RuneUIData" Then
            
            j = 0
            
            'RuneUIData 문서에서 데이터 조회
            strSQL = " SELECT * " & _
                 " FROM [Data$] " & _
                 " WHERE TitleStringKey IN (" & strWhere & ")"
                 
            'OLEDB 연결
            Set objDB = CreateObject("ADODB.Connection")
            objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & strFilePath & ";" & _
                      "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
            
            Set obj = CreateObject("ADODB.Recordset")
            obj.Open strSQL, objDB
            
            ReDim strRuneData(cell.Cells.Count - 1, 1)
            
            '조회된 값이 있는 경우 시트에 표시
            Do Until obj.EOF
            
                If j = 0 Then
                    strRuneData(j, 0) = obj("TitleStringKey")
                    strRuneData(j, 1) = obj.Fields(0)
                    
                    j = j + 1
                    
                '동일한 키에 여러 값들이 존재해서 중복 검토
                ElseIf strRuneData(j - 1, 0) <> obj("TitleStringKey") Then
                    
                    strRuneData(j, 0) = obj("TitleStringKey")
                    strRuneData(j, 1) = obj.Fields(0)
                
                    j = j + 1
                End If

                obj.MoveNext
            Loop
            
            '개체 연결 끊기
            obj.Close
            objDB.Close
            Set obj = Nothing
            Set objDB = Nothing
            
            For k = 0 To j - 1
                Set rngFindCell = cell.Find(strRuneData(k, 0), Lookat:=xlWhole)
                Set rngRuneCell = Sheets("RuneData").UsedRange.Find(strRuneData(k, 1), Lookat:=xlWhole)
                
                rngFindCell.Offset(0, 1) = rngRuneCell.Offset(0, 1).Value
                rngFindCell.Offset(0, 2) = rngFileName(i) '---조회된 파일명 입력
            Next
                  
        Else
            'DATA 시트에서 조건에 맞는 TemplateId, StringId 열의 데이터 추출
            strSQL = " SELECT TemplateId, StringId " & _
                 " FROM [DATA$] " & _
                 " WHERE StringId IN (" & strWhere & ")"
            
                'OLEDB 연결
            Set objDB = CreateObject("ADODB.Connection")
            objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & strFilePath & ";" & _
                      "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
            
            Set obj = CreateObject("ADODB.Recordset")
            obj.Open strSQL, objDB
            
            '조회된 값이 있는 경우 시트에 표시
            Do Until obj.EOF
                
                Set rngFindCell = cell.Find(obj("StringId"), Lookat:=xlWhole)
                rngFindCell.Offset(0, 1) = obj("TemplateId") '---TID값 입력
                rngFindCell.Offset(0, 2) = rngFileName(i) '---조회된 파일명 입력
                
                obj.MoveNext
            Loop
            
            '개체 연결 끊기
            obj.Close
            objDB.Close
            Set obj = Nothing
            Set objDB = Nothing
            
        End If
    Next
    
End Function

'RequestCreateItem InItemType, InTemplateld, InCount, InLevel
Public Sub CheatCreatItem()
    
    Dim InItemType As Variant
    Dim InTemplateId As Variant
    Dim InCount As Variant
    Dim InLevel As Variant
    
    Call SetRange
    
    치트키.ClearContents
    
    For i = 0 To 검색목록.Cells.Count - 1
        
        With 검색목록(i + 1)
        
            InTemplateId = .Offset(0, 1).Value
            
            If .Offset(0, 2).Value = "RangedWeaponData" Or .Offset(0, 2).Value = "AccessoryData" Or .Offset(0, 2).Value = "ReactorData" Then
                InItemType = 2
                
            ElseIf .Offset(0, 2).Value = "ConsumableItemData" Then
                InItemType = 3
                
            ElseIf .Offset(0, 2).Value = "RuneUIData" Then
                InItemType = 4
                
            ElseIf .Offset(0, 2).Value = "CustomizingItemData" Then
                InItemType = 7
            End If
            
            InCount = .Offset(0, 3).Value
            If InCount = 0 Then
                InCount = 1
            End If
            
            InLevel = .Offset(0, 4).Value
            If InLevel = 0 Then
                InLevel = 100
            End If
        
        End With
        
        If InTemplateId = 0 Then
            치트키_시작.Offset(i, 0).Value = "조회된 TID가 존재하지 않습니다."
        Else
        
            치트키_시작.Offset(i, 0).Value = "RequestCreateItem " & InItemType & " " & InTemplateId & " " & _
                                        InCount & " " & InLevel
        End If
    
    Next
    
End Sub

Public Sub ClearKEYList()

    Call SetRange
    
    검색목록.Resize(, 5).ClearContents
    
End Sub

Public Sub ClearCheatList()
    
    Call SetRange
    
    치트키.ClearContents
    
    치트키_시작.Offset(-1, 0).Value = "일괄 입력 희망 시 [메모장 생성] 버튼을 클릭해주세요."
    
End Sub

Public Sub ClearSearchList()
    
    Call SetRange
    
    키목록.Offset(0, -1).Interior.Color = vbWhite
    
End Sub

Public Sub WriteCheat()
    
    Dim path As String
    
    Call SetRange
    
    If IsEmpty(치트키_시작) Then
        MsgBox "생성된 치트키가 없습니다."
        Exit Sub
    End If
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    Open path For Output As #1
        
        Print #1, "<Mag_CreatItem>"
        
        '치트키 영역을 돌면서 반복
        For Each cell In 치트키
            
            '조회된 TID~~ 셀 제외
            If InStr(cell.Value, "조회된") = 0 Then
                Print #1, cell.Value '---작성된 치트 메모장에 출력
            End If
            
        Next
    Close
    
    치트키_시작.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " <Mag_CreatItem>"

End Sub
