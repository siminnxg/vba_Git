Attribute VB_Name = "ModuleFile"
Option Explicit

'ADO로 다른 엑셀 파일을 열지 않고 데이터 가져오기
Sub ADO()
    
    strSQL = " SELECT TemplateId, StringId " & _
                 " FROM [DATA$] " & _
                 " WHERE StringId IN (" & strWhere & ")"
            
    'OLEDB 연결
    Set objDB = CreateObject("ADODB.Connection")
    objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & strFilePath & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    
    Set Obj = CreateObject("ADODB.Recordset")
    Obj.Open strSQL, objDB
    
    '조회된 값이 있는 경우 시트에 표시
    Do Until Obj.EOF
                
        Range("A1").CopyFromRecordset Obj
        
        Obj.MoveNext
    Loop
    
    '개체 연결 끊기
    Obj.Close
    objDB.Close
    Set Obj = Nothing
    Set objDB = Nothing
            
End Sub



Sub 탐색기_폴더_추출()
    
    Dim Selected As Long '---선택한 파일 정보 저장 변수
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "폴더를 선택하세요"
        
        Selected = .Show '---파일 탐색기 열기
        
        '선택된 폴더가 있는 경우 동작
        If Selected = -1 Then
        
            Sheets("etc").Range("H2") = .SelectedItems(1)
        
        '선택된 폴더가 없는 경우 알림
        Else
        
            MsgBox "선택된 폴더가 없습니다."
            
        End If
    End With
    
End Sub


Sub 탐색기_파일_추출()

    '파일 경로가 입력되어 있으면 해당 경로로 지정
    '(잘못된 경로 입력 시 자동으로 무시됨)
    If 파일경로(1) <> "" Then
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = 파일경로(1)

    End If
    
    '파일 탐색기 오픈
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "엑셀파일", "*.xls; *.xlsx; *.xlsm" '---엑셀 형식으로 지정
        .Show
        
        '파일 미 선택 시 종료 처리
        If .SelectedItems.Count = 0 Then
        
            MsgBox "파일을 선택하지 않았습니다."
            Exit Sub
            
        '1개 파일 선택 시 기존 파일명 리스트 하위에 붙여넣기
        ElseIf .SelectedItems.Count = 1 And 파일명.Count < 10 And 파일명(1) <> "" Then
            
            varFileNum = InStrRev(.SelectedItems(1), "\") '---'\' 기준으로 파일경로와 파일명 구분
            파일명(파일명.Count).Offset(1, 0) = Mid(.SelectedItems(1), varFileNum + 1) '---파일명 입력
            파일경로(파일경로.Count).Offset(1, 0) = Left(.SelectedItems(1), varFileNum - 1) '---파일경로 입력
            
            Exit Sub
            
        End If
        
    
End Sub

Sub 파일존재여부()
    
    Path = "파일, 폴더 경로"
    
    If Dir(Path, vbDirectory) = "" Then
    
        MsgBox Path & " 파일은 존재하지 않는 파일입니다." & vbCrLf & vbCrLf & _
                "경로를 확인해주세요."
        
    End If

End Sub

Sub GetObject()
    
    
    Path = "파일경로 & 이름"
    
    Set Obj = GetObject(Path)
    
    Set wb = Workbooks(Dir(Path))
    
    shtname = wb.Sheets(1).Name '---첫번째 시트명 추출
    
    Set WS = wb.Sheets(shtname)
    
    MsgBox Application.WorksheetFunction.CountIf(WS.UsedRange, 검색어) '---설정한 시트에 검색어와 일치하는 셀 개수가 몇개인지 확인
    
    MsgBox Application.WorksheetFunction.CountIf(WS.UsedRange, "*" & 검색어 & "*")
    
    
    
    
    
    Set rngFind = WS.UsedRange.Find(what:=검색어, lookat:=xlPart) '부분 일치, 정확히 일치 : xlWhole
    
End Sub

Sub 쿼리_호출()
    
    Path = "파일경로 & 이름"
    
    '입력된 경로를 바탕으로 파일 불러오기
        ActiveWorkbook.Queries.Add Name:=preset, _
        Formula:="let Source = Excel.Workbook(File.Contents(""" & Path & """), null, true), #""" & _
                sheet_name & "_Sheet"" = Source{[Item=""" & sheet_name & """, Kind=""Sheet""]}[Data], " & _
                "FilteredData = Table.PromoteHeaders(#""" & sheet_name & "_Sheet"") " & _
        "in FilteredData"
        
        '연결된 쿼리 데이터 가져오기
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & preset & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & preset & "]")
            .Refresh BackgroundQuery:=False
        End With
    
End Sub
