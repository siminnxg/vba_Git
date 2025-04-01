Attribute VB_Name = "ModuleMain"
'=====================================================================
'��ũ�� : LoadFile
'��� ��Ʈ : Home ��Ʈ
'���� : �ҷ����� ��ư ����, �ԷµǾ� �ִ� ���� ������ �������� ������ ȣ��
'=====================================================================
Public Sub LoadFile()

    '###���� ����###
    
    Dim varTmp As Variant '---�ӽ� ����
    Dim varFileName As Variant '---Ȯ���� ������ ���ϸ�
    
    '###���� ����###
    
    '������ �߻��ϸ� ����
    On Error GoTo exit_error
        
    Call UpdateStart
    Call SetRange
    
    '����� �Է� ���� ���� �� �˸� ǥ��
    If File_adr = "" Or File_name = "" Or sheet_name = "" Then
    
        MsgBox "���� ������ ��� �Է����ּ���." & vbCrLf & "(���� ���, �̸�, ��Ʈ)"
        GoTo exit_sub

    End If
    
    '������ ���� �� �ӽ� �̸� ����
    If preset = "" Or preset = "������" Then
        
        preset = CheckPresetName
        �����¸� = preset
    
    End If
    
    '������ �̸����� ��Ʈ, ���� �̹� �����Ǿ� �ִٸ� ����
    If CheckQuery = 1 Then
        
        MsgBox ("������ �����¸��� �����մϴ�.")
        GoTo exit_sub
    
    '�Է��� ��ο� ���� ���� ���� üũ
    ElseIf CheckFile(File_adr & "\" & File_name) = False Then
        
        MsgBox (File_adr & " ��ο� " & File_name & " ������ �������� �ʽ��ϴ�.")
        GoTo exit_sub
    
    '������ ���� ������ �ƴ� ��� ó��
    ElseIf InStr(File_name, ".xl") = 0 Then
     
         MsgBox "������ ���� ������ �ƴմϴ�."
         GoTo exit_sub
         
    Else
        'Search ��Ʈ ǥ�� �� ����
        Call HideSearchSht(False)
        Sheets("Search").Select
        
        '���� �˻� ������ ���� �� �ʱ�ȭ
        Call SaveSearch
        Call ClearHomeData
                
        '������ �̸����� ��Ʈ ����
        ActiveWorkbook.Worksheets.Add after:=Sheets("Search")
        ActiveSheet.Name = preset
        
        '�Էµ� ��θ� �������� ���� �ҷ�����
        ActiveWorkbook.Queries.Add Name:=preset, _
        Formula:="let Source = Excel.Workbook(File.Contents(""" & File_adr & "\" & File_name & """), null, true), #""" & _
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
        
        Call SavePreset '---������ ����
        
        Sheets(preset).Visible = 2 '---��Ʈ ǥ��
    
    End If
        
    Sheets("Search").Select
    
    '���� ȣ�� �� �� ���� ���� ���� ǥ��
    Call HideHomeCategory(False)
    
    'ī�װ� ����Ʈ ȣ��
    Call SearchCategory(preset)
    
    'ī�װ� ��ü ���� �� �߰�
    If SelectAllCategory = 0 Then
        
        varCheckUpdate = Empty
        Call SelectAllCategory
        Call Button_AddCategory
        
    End If

'���� ó��
exit_sub:
    
    Call UpdateEnd
    Range("A1").Select
    Exit Sub
    
'���� �߻� ó��
exit_error:
    
    '---��Ʈ�� ���� �� ó�� (��Ʈ�� ���� üũ, ������ ���� üũ)
    If Left(Err.Description, 6) = "�Է��� ��Ʈ" Or Err.Number = -2147024809 Then
        
        '---��Ʈ ���� �� �ý��� ���� �̳���
        Application.DisplayAlerts = False
        
        '---������ ��Ʈ ����
        ActiveSheet.Delete
        Sheets("Home").Select
        
        Application.DisplayAlerts = True
        
        '---�˸� ǥ��
        MsgBox "�����¸��� �������ּ���" & vbCrLf & vbCrLf & Err.Description
        GoTo exit_sub
        
    End If
    
    '---�� �� ���� ó��
    MsgBox ("������ �߻��߽��ϴ�. ���� �ڵ� : " & Err.Number & vbCrLf & Err.Description)
    
    Call UpdateEnd
    Range("A1").Select
    
End Sub

'=====================================================================
'��ũ�� : SearchFile
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ���� �˻� ��ư Ŭ�� �� ����, ���� Ž���� ���� �� ����ڰ� ������ ������ ���, �̸� ȣ��
'=====================================================================
Public Sub SearchFile()
    
    '###���� ����###
    
    Dim sel_File As Variant '---������ ���ϸ� ���� ����
    Dim wb As Object '---������ ���� ������Ʈ ���� ����
    Dim sheet_count As Variant '---�ҷ��� ������ ��Ʈ ���� ���� ����
    Dim file_open_check As String '---�ҷ��� ���� ���� ���� Ȯ�� ����
    Dim var() As String '---�ӽ� ��� ����
    
    '###���� ����###
    
    '--- sub �������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call UpdateStart
    Call SetRange
    
    '---���� �ּ� ��������
    sel_File = File_adr & "\" & File_name
    
    '---etc��Ʈ �� ��Ʈ�� ���� ���� �ʱ�ȭ
    Sheets("etc").Range("A:A").Clear
    
    '���� ��ΰ� �ԷµǾ� ������ �ش� ��η� ����
    '(�߸��� ��� �Է� �� �ڵ����� ����)
    If File_adr <> "" Then
     
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = File_adr
         
    End If
    
    '���� Ž���� ����
     With Application.FileDialog(msoFileDialogFilePicker)
         .Filters.Add "��������", "*.xls; *.xlsx; *.xlsm" '---���� �������� ����
         .Show
         
         '---���� �� ���� �� ���� ó��
         If .SelectedItems.Count = 0 Then
         
             MsgBox "������ �������� �ʾҽ��ϴ�."
             GoTo exit_sub
             
         End If
         
         '---������ ������ sel_File ������ ����
         sel_File = .SelectedItems(1)
         
     End With
     
     '---������ ������ ���� ������ �ƴ� ��� ó��
     If InStr(sel_File, ".xl") = 0 Then
     
         MsgBox "���� ������ �������ּ���."
         GoTo exit_sub
         
     End If
                    
    '---���� ��� �� �̸� �и� �� ����
    max = InStrRev(sel_File, "\")
    ���ϰ��.Value = Left(sel_File, max - 1)
    ���ϸ�.Value = Mid(sel_File, max + 1)
    
    '---sub : ���� ���� ��ȣ��
    Call LoadFileInfo
           
    '---function : ����ڰ� ������ ������ �̹� �����ִ��� Ȯ��
    file_open_check = CheckFileOpen(���ϸ�.Value)
               
    '---������ ������ ��Ʈ�� �ҷ�����
    Set wb = GetObject(sel_File)
    
    '---�ش� ������ ��Ʈ ���� Ȯ��
    sheet_count = wb.Sheets.Count
    
    '---��Ʈ ������ŭ ����
    For n = 1 To sheet_count
    
        Sheets("etc").Range("a" & n) = wb.Sheets(n).Name
        
    Next
    
    '---sheet_list �̸� ���� ������
    ThisWorkbook.Names("sheet_list").RefersTo = Sheets("etc").Range("A1", Sheets("etc").Cells(n - 1, 1))
    
    '---������ �������� �ʴ� �����̶�� ���� ����
    If file_open_check = True Then
        
        wb.Close
        
    End If
    
    '---��Ʈ ��� ��� �ٿ����� ǥ��
    With ��Ʈ��.Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=sheet_list"
    End With
    
    '---ù��° ��Ʈ�� �⺻������ ǥ��
    ��Ʈ��.Value = Sheets("etc").Range("A1").Value
    
'---���� ó��
exit_sub:

    Call UpdateEnd
    Range("A1").Select
    
End Sub

'=====================================================================
'��ũ�� : SearchCategory
'��� ��Ʈ : Home ��Ʈ
'���� : ���� �˻� ��ư Ŭ�� �� ����, ���� Ž���� ���� �� ����ڰ� ������ ������ ���, �̸� ȣ��
'=====================================================================
Public Sub SearchCategory(sheet_name)
    
    '###���� ����###
        
    Dim category() As Variant   '---ī���� ������ ������ �迭
    Dim category_row As Range   '---�ҷ��� ������ ��Ʈ���� ù �� ���� ����
    Dim Target As Range         '---���� ���õ� ������ �� ���� ����
    Dim category_range As Range '---ī�װ� ���� ����
    
    '###���� ����###
    
    ����������.Value = sheet_name
        
    Call HideSearchSht(False)
    
    '---ī�װ� ����Ʈ ���� �ʱ�ȭ
    �����.Clear
    
    '---������ ��Ʈ�� ù��° �� ���� ����
    Set category_list = Sheets(sheet_name).ListObjects(1).HeaderRowRange
    
    '---�迭 ũ�� ������
    ReDim category(category_list.Columns.Count - 1)
    
    '---�迭�� ù���� �� �� ���� ����
    For i = 0 To category_list.Columns.Count - 1
    
        category(i) = Sheets(sheet_name).Range("A1").Offset(0, i)
        
    Next
    
    '---Home ��Ʈ�� �迭 �� �ѷ��ֱ�
    For i = 0 To UBound(category)
    
        ����������.Offset(i + 1, 0).Value = category(i)
    
    Next
    
    '---ī�װ� ���� �� �Ҵ�
    Call SetRange
    
    '---ī�װ� ���� �׵θ� ����
    With �����
            
            .Borders.LineStyle = 1
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 1
        
    End With
    
    '---���� �� ���� �� �ʱ�ȭ
    ������.ClearContents
        
End Sub

'=====================================================================
'��ũ�� : AddCategory
'��� ��Ʈ : Home ��Ʈ
'���� : ���� �˻� ��ư Ŭ�� �� ����, ���� Ž���� ���� �� ����ڰ� ������ ������ ���, �̸� ȣ��
'=====================================================================
Public Function AddCategory()

    'AddCategory = 1 : ���� �߻�
    'AddCategory = 0 : ���� ����
    
    '###���� ����###
    
    Dim rngSel As Range '---���� ���õ� ��
    Dim select_category As Range
    
    Dim category_count As Variant
    Dim search_row As Variant
    Dim varCol As Variant
    
    '###���� ����###
    
    '---�������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call SetColor
    Call SetRange
    
    '---function : ī�װ��� �������� ������ ���� ����
    If CheckCategory = 1 Then
    
        Range("notice") = "ī�װ� ����Ʈ�� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        AddCategory = 1
        Exit Function
        
    End If
       
    Range("notice") = ""
    
    category_count = 0
    
    If �˻�Ű����_���� <> Empty Then

        'Call ResetSearch '���� ���� Ȯ�� �� �ʱ�ȭ

        Range("DATA").Clear '�˻� ���� �ʱ�ȭ

    End If
    
    Sheets(����������.Value).UsedRange.EntireColumn.Hidden = False '---�� ����� ���
    
    '���õ� ī�װ� üũ
    For i = 1 To �����.Rows.Count
        
        '��ȸ���� �� ����
        Set rngSel = �����_����.Offset(i - 1, 0)
        
        '�������� ���� ���� üũ
        If Not rngSel.Interior.Color = vbWhite Then
            
            �˻�Ű����_����.Offset(0, category_count).Value = rngSel.Value '---���õ� �� ������ �� �� �� �߰�
                        
            �˻���_����.Offset(0, category_count).Interior.Color = colorUserInput '---���õ� �� ���� �� ����
                        
            category_count = category_count + 1 '---���õ� �� ���� üũ
            
            'select_category ������ ���õ� �� �ּ� ����
            If select_category Is Nothing Then
                Set select_category = rngSel
                
            Else
                Set select_category = Union(select_category, rngSel)
                
            End If
        
        '�� ���� ������ ��� ������ ��Ʈ���� �� �����
        Else
            With Sheets(����������.Value)
            
                varCol = .ListObjects(1).ListColumns(rngSel.Value).Index
                .Columns(varCol).Hidden = True
                
            End With
        End If
    Next
    
    '---���õ� ī�װ��� ���� ���
    If category_count = 0 Then
    
        Range("notice") = "���õ� ī�װ��� �����ϴ�."
        Range("notice").Font.Color = vbRed
        
        AddCategory = 1
        
    Else
        
        Call LoadSearch
        Call PasteData
        Call SetRange
        
        'ī�װ� �߰��� �� �ʺ� ����
        �˻�Ű����.EntireColumn.AutoFit
        
        '�˻� ���� �ؽ�Ʈ Ÿ������ ����
        Range(�˻���_����, �˻���_����.Offset(0, category_count - 1)).NumberFormatLocal = "@"
        
        AddCategory = 0
    End If
    
    With Sheets("etc")

        '---etc ��Ʈ������ ������ ��ġ ã��
        search_row = .Range("preset_list").Find(what:=����������.Value, lookat:=xlWhole).Row

        '---���õ� �� �ּ� �����¿� �����ϱ�
        If Not select_category Is Nothing Then

            .Cells(search_row, 6) = select_category.Address

        Else

            .Cells(search_row, 6).Clear

        End If

    End With

End Function

'=====================================================================
'��ũ�� : ResetCategory
'��� ��Ʈ : Home ��Ʈ
'���� : �� ���� ���� ���� ���� ��ư ����, �� ����Ʈ ��ü �� ���� ���ֱ�
'=====================================================================
Public Sub ResetCategory()

    '---�������� ����ϴ� ���� ��ġ ȣ��
    Call SetRange
    
    '---function : ǥ�õ� �� �������� ������ ���� ����
    If CheckCategory = 1 Then
    
        Range("notice") = "ī�װ� ����Ʈ�� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        Exit Sub
    End If
    
    '---ī�װ� ���� ���� �ʱ�ȭ
    �����.Interior.Color = vbWhite
    
End Sub

'=====================================================================
'��ũ�� : SelectAllCategory
'��� ��Ʈ : Home ��Ʈ
'���� : �� ���� ���� ��ü ���� ��ư ����, �� ����Ʈ ��ü �� ���� ����
'=====================================================================
Public Function SelectAllCategory()
    
    'SelectAllCategory = 1 : ǥ�õ� ���� ����
    
    '---�������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call SetColor
    Call SetRange
    
    '---function : ǥ�õ� ���� �������� ������ ���� ����
    If CheckCategory = 1 Then
    
        Range("notice") = "ī�װ� ����Ʈ�� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        SelectAllCategory = 1
        Exit Function
        
    End If
    
    '---�� ��ü ���� �� ���� ����
    �����.Interior.Color = colorCategorySel
    SelectAllCategory = 0

End Function
