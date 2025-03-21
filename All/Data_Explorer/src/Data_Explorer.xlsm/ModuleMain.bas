Attribute VB_Name = "ModuleMain"
'=====================================================================
'��ũ�� : LoadFile
'��� ��Ʈ : Home ��Ʈ
'���� : �ҷ����� ��ư ����, �ԷµǾ� �ִ� ���� ������ �������� ������ ȣ��
'=====================================================================
Public Sub LoadFile()

    '###���� ����###
    
    Dim var As Variant '---�ӽ� ����
    Dim File_name_val As Variant '---Ȯ���� ������ ���ϸ�
    
    '###���� ����###
    
    '---������ �߻��ϸ� ����
    On Error GoTo exit_error
    
    '---sub : ȭ�� ������Ʈ ����
    Call UpdateStart
    
    '---sub : �������� ����ϴ� ���� ��ġ ȣ��
    Call SetRange
    
    '---����� �Է� ���� ���� �� �˸� ǥ��
    If File_adr = "" Or File_name = "" Or sheet_name = "" Then
    
        MsgBox "���� ������ ��� �Է����ּ���." & vbCrLf & "(���� ���, �̸�, ��Ʈ)"
        GoTo exit_sub

    End If
    
    '---function : ������ ���� �� �ӽ� �̸� ����
    If preset = "" Or preset = "������" Then
        
        preset = CheckPresetName
        �����¸� = preset
    
    End If
    
    '---���� �̸����� Ȯ����(.xl~) �и�
    var = Split(File_name, ".")
    File_name_val = var(0)
    
    '---function : ������ �̸����� ��Ʈ, ���� �̹� �����Ǿ� �ִٸ� ����
    If CheckQuery = 1 Then
        
        MsgBox ("������ �����¸��� �����մϴ�.")
        GoTo exit_sub
    
    '---function : �Է��� ��ο� ���� ���� ���� üũ
    ElseIf CheckFile(File_adr & "\" & File_name) = False Then
        
        MsgBox (File_adr & " ��ο� " & File_name & " ������ �������� �ʽ��ϴ�.")
        GoTo exit_sub
    
    '---������ ���� ������ �ƴ� ��� ó��
    ElseIf InStr(File_name, ".xl") = 0 Then
     
         MsgBox "������ ���� ������ �ƴմϴ�."
         GoTo exit_sub
         
    Else
    
        '---sub : �� ����, �˻� ���� �ʱ�ȭ
        Call ClearHomeData
                
        '---������ �̸����� ��Ʈ ����
        ActiveWorkbook.Worksheets.Add after:=Sheets("Home")
        ActiveSheet.Name = preset
        
        '---�Էµ� ��θ� �������� ���� �ҷ�����
        ActiveWorkbook.Queries.Add Name:=preset, _
        Formula:="let Source = Excel.Workbook(File.Contents(""" & File_adr & "\" & File_name & """), null, true), #""" & _
                sheet_name & "_Sheet"" = Source{[Item=""" & sheet_name & """, Kind=""Sheet""]}[Data], " & _
                "FilteredData = Table.PromoteHeaders(#""" & sheet_name & "_Sheet"") " & _
        "in FilteredData"

                
'                "��翭���� = Table.TransformColumnTypes(FilteredData, " & _
'        "List.Transform(Table.ColumnNames(FilteredData), each {_, type text})) " & _
'        "in ��翭����"
        
        '---����� ���� ������ ��������
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & preset & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & preset & "]")
            .Refresh BackgroundQuery:=False
        End With
        
        '---sub : ������ ����
        Call preset_save
        
    End If
        
    Sheets("Home").Select
    
    '---���� ȣ�� �� �� ���� ���� ���� ǥ��
    Sheets("Home").Columns("G:H").Hidden = False
    Sheets("Home").Shapes("Pic_Open").Visible = False
    Sheets("Home").Shapes("PIC_Close").Visible = True
    
    '---sub : ī�װ� ����Ʈ ȣ��
    Call SearchCategory(preset)
    
    '---function : ī�װ� ��ü ���� �� �߰�
    If SelectAllCategory = 0 Then
        
        varCheckUpdate = Empty
        Call SelectAllCategory
        Call Button_AddCategory
        
    End If

'---���� ó��
exit_sub:
    
    Call UpdateEnd
    Range("A1").Select
    Exit Sub
    
'---���� �߻� ó��
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
    
    '---���� ��ΰ� �ԷµǾ� ������ �ش� ��η� ����
    '---(�߸��� ��� �Է� �� �ڵ����� ���õ�)
    If File_adr <> "" Then
     
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = File_adr
         
    End If
    
    '---���� Ž���� ����
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
    
    '---�� ����, �˻� ���� �����
    Call HideHomeData
    
    '---ī�װ� ����Ʈ ���� �ʱ�ȭ
    �����.Clear
    
    '---������ ��Ʈ�� ù��° �� ���� ����
    Set category_list = Sheets(sheet_name).Range("A1", Sheets(sheet_name).Range("A1").End(xlToRight))
    
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
    
    Dim now_cell As Range
    Dim category_count As Variant
    Dim select_category As Range
    Dim search_row As Variant
    
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
        
        '---sub : ���� ���� Ȯ�� �� �ʱ�ȭ
        Call ResetSearch
            
        '---�˻�, ī�װ� ���� �ʱ�ȭ
        Range("DATA").Clear
        
    End If
    
    '---���õ� ī�װ� üũ
    For i = 1 To �����.Rows.Count
        
        '---��ȸ���� �� ����
        Set now_cell = �����_����.Offset(i - 1, 0)
        
        '---�������� ���� ���� üũ
        If Not now_cell.Interior.Color = vbWhite Then
            
            '---���õ� �� ������ �� �� �� �߰�
            �˻�Ű����_����.Offset(0, category_count).Value = now_cell.Value
            
            '---���õ� �� ���� �� ����
            �˻���_����.Offset(0, category_count).Interior.Color = colorUserInput
            
            '---�� ���� üũ
            category_count = category_count + 1
            
            '---select_category ������ ���õ� �� �ּ� ����
            If select_category Is Nothing Then
            
                Set select_category = now_cell
                
            Else
            
                Set select_category = Union(select_category, now_cell)
                
            End If
        End If
    Next
    
    '---���õ� ī�װ��� ���� ���
    If category_count = 0 Then
    
        Range("notice") = "���õ� ī�װ��� �����ϴ�."
        Range("notice").Font.Color = vbRed
        
        AddCategory = 1
        
    Else
    
        '---ī�װ� �߰��� �� �ʺ� ����
        Range(�˻�Ű����_����, �˻�Ű����_����.Offset(0, category_count - 1)).EntireColumn.AutoFit
        
        '---�˻� ���� �ؽ�Ʈ Ÿ������ ����
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

'=====================================================================
'��ũ�� : ResetSearch
'��� ��Ʈ : Home ��Ʈ
'���� : �˻� ���� �˻� �ʱ�ȭ ��ư, �˻����̴� ���� �ʱ�ȭ �� ������ ��Ʈ ���� ����
'=====================================================================
Public Function ResetSearch()
    
    'ResetSearch = 1 : ���õ� ���� ����
    'ResetSearch = 0 : ���� ����
        
    '---�������� ����ϴ� ���� ��ġ ȣ��
    Call SetRange
    
    '---�˻��� �ʱ�ȭ
    Range(�˻���_����, �˻�Ű����_����.End(xlToRight).Offset(-1, 0)).ClearContents
    
    '---���õ� �� �� ���� �ʱ�ȭ
    Range(�˻�Ű����_����, �˻�Ű����_����.End(xlToRight)).ClearFormats
    
    '---���õ� ���� �������� ������ ���� ����
    If �˻�Ű����_���� = "" Then
    
        Range("notice") = "���õ� ī�װ��� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        ResetSearch = 1
        
        Exit Function
        
    End If
    
    '---������ ��Ʈ ������ ���Ͱ� �ɷ��ִٸ� ����
    If Sheets(CStr(����������.Value)).AutoFilter.FilterMode = True Then

        Sheets(����������.Value).ShowAllData

    End If
    
    ResetSearch = 0
    
End Function
