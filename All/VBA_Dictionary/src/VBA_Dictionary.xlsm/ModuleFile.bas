Attribute VB_Name = "ModuleFile"
Option Explicit

'ADO�� �ٸ� ���� ������ ���� �ʰ� ������ ��������
Sub ADO()
    
    strSQL = " SELECT TemplateId, StringId " & _
                 " FROM [DATA$] " & _
                 " WHERE StringId IN (" & strWhere & ")"
            
    'OLEDB ����
    Set objDB = CreateObject("ADODB.Connection")
    objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & strFilePath & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    
    Set Obj = CreateObject("ADODB.Recordset")
    Obj.Open strSQL, objDB
    
    '��ȸ�� ���� �ִ� ��� ��Ʈ�� ǥ��
    Do Until Obj.EOF
                
        Range("A1").CopyFromRecordset Obj
        
        Obj.MoveNext
    Loop
    
    '��ü ���� ����
    Obj.Close
    objDB.Close
    Set Obj = Nothing
    Set objDB = Nothing
            
End Sub



Sub Ž����_����_����()
    
    Dim Selected As Long '---������ ���� ���� ���� ����
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "������ �����ϼ���"
        
        Selected = .Show '---���� Ž���� ����
        
        '���õ� ������ �ִ� ��� ����
        If Selected = -1 Then
        
            Sheets("etc").Range("H2") = .SelectedItems(1)
        
        '���õ� ������ ���� ��� �˸�
        Else
        
            MsgBox "���õ� ������ �����ϴ�."
            
        End If
    End With
    
End Sub


Sub Ž����_����_����()

    '���� ��ΰ� �ԷµǾ� ������ �ش� ��η� ����
    '(�߸��� ��� �Է� �� �ڵ����� ���õ�)
    If ���ϰ��(1) <> "" Then
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = ���ϰ��(1)

    End If
    
    '���� Ž���� ����
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "��������", "*.xls; *.xlsx; *.xlsm" '---���� �������� ����
        .Show
        
        '���� �� ���� �� ���� ó��
        If .SelectedItems.Count = 0 Then
        
            MsgBox "������ �������� �ʾҽ��ϴ�."
            Exit Sub
            
        '1�� ���� ���� �� ���� ���ϸ� ����Ʈ ������ �ٿ��ֱ�
        ElseIf .SelectedItems.Count = 1 And ���ϸ�.Count < 10 And ���ϸ�(1) <> "" Then
            
            varFileNum = InStrRev(.SelectedItems(1), "\") '---'\' �������� ���ϰ�ο� ���ϸ� ����
            ���ϸ�(���ϸ�.Count).Offset(1, 0) = Mid(.SelectedItems(1), varFileNum + 1) '---���ϸ� �Է�
            ���ϰ��(���ϰ��.Count).Offset(1, 0) = Left(.SelectedItems(1), varFileNum - 1) '---���ϰ�� �Է�
            
            Exit Sub
            
        End If
        
    
End Sub

Sub �������翩��()
    
    Path = "����, ���� ���"
    
    If Dir(Path, vbDirectory) = "" Then
    
        MsgBox Path & " ������ �������� �ʴ� �����Դϴ�." & vbCrLf & vbCrLf & _
                "��θ� Ȯ�����ּ���."
        
    End If

End Sub

Sub GetObject()
    
    
    Path = "���ϰ�� & �̸�"
    
    Set Obj = GetObject(Path)
    
    Set wb = Workbooks(Dir(Path))
    
    shtname = wb.Sheets(1).Name '---ù��° ��Ʈ�� ����
    
    Set WS = wb.Sheets(shtname)
    
    MsgBox Application.WorksheetFunction.CountIf(WS.UsedRange, �˻���) '---������ ��Ʈ�� �˻���� ��ġ�ϴ� �� ������ ����� Ȯ��
    
    MsgBox Application.WorksheetFunction.CountIf(WS.UsedRange, "*" & �˻��� & "*")
    
    
    
    
    
    Set rngFind = WS.UsedRange.Find(what:=�˻���, lookat:=xlPart) '�κ� ��ġ, ��Ȯ�� ��ġ : xlWhole
    
End Sub

Sub ����_ȣ��()
    
    Path = "���ϰ�� & �̸�"
    
    '�Էµ� ��θ� �������� ���� �ҷ�����
        ActiveWorkbook.Queries.Add Name:=preset, _
        Formula:="let Source = Excel.Workbook(File.Contents(""" & Path & """), null, true), #""" & _
                sheet_name & "_Sheet"" = Source{[Item=""" & sheet_name & """, Kind=""Sheet""]}[Data], " & _
                "FilteredData = Table.PromoteHeaders(#""" & sheet_name & "_Sheet"") " & _
        "in FilteredData"
        
        '����� ���� ������ ��������
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & preset & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & preset & "]")
            .Refresh BackgroundQuery:=False
        End With
    
End Sub
