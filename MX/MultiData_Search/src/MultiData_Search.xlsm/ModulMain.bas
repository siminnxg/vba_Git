Attribute VB_Name = "ModulMain"
Option Explicit

'=====================================================================
'��ũ�� : SearchData
'��� ��Ʈ : Main ��Ʈ
'���� : ����ڰ� �Է��� ������ ���ϵ鿡�� �˻�� ã�� �ش� ���� �����͸� ��� �����ɴϴ�.
'=====================================================================
Public Sub SearchData()
    
    '###���� ����###
    
    'ȣ�� ���� ���� ����
    Dim Obj As Object
    Dim wb As Workbook
    Dim WS As Worksheet
    
    Dim strFile As String '---���� ��� & �̸�
    Dim strSheet As String '---��Ʈ��
    
    'ȣ���� ���� ������ ���� ����
    Dim rngWS_Search As Range '---�˻� ��ġ ����
    Dim rngWS_Result As Range '---�˻� ��� ����
    Dim varWS_Col As Variant '---�˻����� ��Ʈ�� �� ����
    
    Dim strSearchStart As String '---ù��° �˻��� �ּ� ����
    Dim varResultCount As Variant '---�˻��� ����
    
    
    '###���� ����###
    On Error Resume Next
        
    Call UpdateStart '---ȭ�� ������Ʈ ����
    Call SetRange '---�� ��� ���� ����
    
    '����� �Է� ������ ���� ��� ó��
    If CheckUserData = True Then
        
        GoTo exit_sub
    
    '��ο� ���� ���� ��� ó��
    ElseIf CheckFile() = True Then
                
        GoTo exit_sub
        
    End If
    
    varResultCount = 0
    
    '�˻��� ���� Ȯ��
    For i = 1 To ���ϸ�.count
        
        strFile = ���ϰ��(i) & "\" & ���ϸ�(i)
        strSheet = ���ϸ�(i).Offset(0, 1).Value
    
        '���� ȣ��
        If CheckFileOpen(strFile) = False Then
                
            Set Obj = GetObject(strFile)
        
        End If
        
        Set wb = Workbooks(Dir(strFile))
        
        Call ObjectList(strFile)
        
        '��Ʈ�� ���� �� ù��° ��Ʈ �⺻������ ����
        If strSheet = "" Then
        
            strSheet = wb.Sheets(1).Name
            ���ϸ�(i).Offset(0, 1) = strSheet
            
        End If
        
        '���Ͽ� �ش� ��Ʈ ���� ��� ����
        If CheckSheet(wb, strSheet) = True Then
            MsgBox ���ϸ�(i) & " ���Ͽ� " & strSheet & " ��Ʈ�� �������� �ʽ��ϴ�."
            
            Obj.Close '---ȣ��� ���� �ݱ�
            GoTo exit_sub
            
        End If
        
        Set WS = wb.Sheets(strSheet)
                
        'ȣ��� ���� ������ �˻� ���� üũ
        If �˻���.Offset(0, 1) = "����" Then
            varResultCount = varResultCount + Application.WorksheetFunction.CountIf(WS.UsedRange, "*" & �˻��� & "*") '---�˻� �ɼ� ����
            
        Else
            varResultCount = varResultCount + Application.WorksheetFunction.CountIf(WS.UsedRange, �˻���) '---�˻� �ɼ� ��ġ
            
        End If
    Next
    
    '�˻� ��� 1���� �̻� �� ����
    If varResultCount > 10000 Then
        
        MsgBox "�˻��� ����� " & Format(varResultCount, "0,000") & "�� �Դϴ�. " & _
                vbCrLf & "�����Ͱ� ���� ��ȸ�� ���� �ð��� �ҿ�˴ϴ�." & _
                vbCrLf & vbCrLf & "�ڼ��� �˻�� �Է����ּ���."
                
            GoTo exit_sub
    
    '�˻��� ����� ���� ��� ����
    ElseIf varResultCount = 0 Then
        
        MsgBox "�˻� ����� �����ϴ�."
        GoTo exit_sub
        
    End If
    
    '�˻� ��� ǥ�õǴ� 'DATA' ���� �ʱ�ȭ
    Range("DATA").Clear
    ThisWorkbook.Names("DATA").RefersTo = �˻����
    
    '�Էµ� ���� ������ŭ �ݺ�
    For i = 1 To ���ϸ�.count
    
        strFile = ���ϰ��(i) & "\" & ���ϸ�(i)
        strSheet = ���ϸ�(i).Offset(0, 1).Value
        
        varResultCount = 0 '---�˻� ���� �ʱ�ȭ
        
        '�˻��� ���� ����
        Set wb = Workbooks(Dir(strFile))
        Set WS = wb.Sheets(strSheet)
        
        '�˻�
        If �˻���.Offset(0, 1) = "����" Then
            
            Set rngWS_Search = WS.UsedRange.Find(what:=�˻���, lookat:=xlPart)
            
        Else
            
            Set rngWS_Search = WS.UsedRange.Find(what:=�˻���, lookat:=xlWhole)
            
        End If
        
        '�˻��� ���� �����ϴ� ���
        If Not rngWS_Search Is Nothing Then
                        
            strSearchStart = rngWS_Search.Address '---ó�� �˻��� ��ġ ����

            varWS_Col = WS.UsedRange.Columns.count '---�˻��� ���Ͽ��� ������� �� ���� üũ
            
            '�Ӹ��� �� ����
            If �Ӹ���(i) = "" Then
                Set rngWS_Result = WS.UsedRange.Rows(1)
                
            Else
                Set rngWS_Result = WS.UsedRange.Rows(�Ӹ���(i))
                
            End If
            
            Set rngWS_Result = Union(rngWS_Result, WS.UsedRange.Rows(rngWS_Search.Row)) '---�˻��� ���� ������ �߰�
            
            '���� ������ ���� �˻�
            Do
            
                Set rngWS_Search = WS.UsedRange.FindNext(rngWS_Search) '---�˻�

                Set rngWS_Result = Union(rngWS_Result, WS.UsedRange.Rows(rngWS_Search.Row)) '---�˻��� ���� ������ �߰�
                    
            Loop While Not rngWS_Search Is Nothing And strSearchStart <> rngWS_Search.Address '---�˻� ������ ���ų� ù��° �ּҷ� ���ƿ� ��� ����
            
            '�˻� ��� �ٿ��ֱ�
            �˻���� = ���ϸ�(i) '---ù��° ���� �˻��� ���ϸ� ǥ��
            rngWS_Result.Copy Destination:=�˻����.Offset(0, 1) '---���� ���� �ٿ��ֱ�
            
            '�˻� �� ������� �˻��Ǿ� ������ �࿡ ���� �˻� ����� �����ϴ� ��� �˻� ���� -1
            For Each rngTemp In rngWS_Result.Areas
                
                varResultCount = varResultCount + rngTemp.Rows.count
                
            Next rngTemp
        
            ThisWorkbook.Names("DATA").RefersTo = Range(Range("DATA"), �˻����.Offset(varResultCount, varWS_Col)) '---'DATA' ���� ������
            
            With Range(�˻����.Offset(0, 1), �˻����.Offset(varResultCount - 1, varWS_Col))
                
                '�� �׵θ�
                .Borders(xlLeft).LineStyle = xlContinuous
                .Borders(xlRight).LineStyle = xlContinuous
                .Borders(xlTop).LineStyle = xlContinuous
                .Borders(xlBottom).LineStyle = xlContinuous
                
            End With

            Set �˻���� = �˻����.Offset(varResultCount + 1, 0) '---'�˻����' ���� ������
            
            Application.GoTo reference:=�˻���.Offset(-2, -1), Scroll:=True  ' ���ϴ� ���� �̵� �� ��ũ��
            
        End If
    Next

'���� ó��
exit_sub:
    Call UpdateEnd

End Sub

'=====================================================================
'��ũ�� : CloseFile
'��� ��Ʈ : etc ��Ʈ
'���� : ����ڰ� �Է��� ������ ���ϵ鿡�� �˻�� ã�� �ش� ���� �����͸� ��� �����ɴϴ�.
'=====================================================================
Public Sub CloseFile()
    
    Dim wb As Workbook
    Dim count As Variant
    
    On Error Resume Next
    
    '���� ��ȣ�� ������ ��� ����
    If Range("������Ʈ")(1) = "" Then
        
        Exit Sub
    End If
        
    Call SetRange '---�� ��� ���� ����
    
    'ȣ��� ���ϵ� �ݱ�
    For i = 1 To Range("������Ʈ").count
        
        Set wb = Workbooks(Dir(Range("������Ʈ")(i)))
        wb.Close
        count = 1
        
    Next
    
    If count = 1 Then
            
        ' '������Ʈ' ���� ������ �ʱ�ȭ
        Range(Range("������Ʈ"), Range("������Ʈ").Offset(0, 2)).Clear
        ThisWorkbook.Names("������Ʈ").RefersTo = ������Ʈ
        
    End If
    
End Sub

'=====================================================================
'��ũ�� : OpenFile
'���� : GetObject�� ȣ���� ���ϵ��� ��� ȭ�鿡 ����ݴϴ�.
'=====================================================================
Public Sub OpenFile()

    Dim wb As Workbook
    
    On Error Resume Next
    
    If Range("������Ʈ").Cells(1) = "" Then
        Exit Sub
    End If
    
    Call UpdateStart
    
    For i = 1 To Range("������Ʈ").count
        Set wb = Workbooks(Dir(Range("������Ʈ").Cells(i)))
        
        wb.IsAddin = True
        wb.IsAddin = False
        ThisWorkbook.Activate
        
    Next
    
    Call UpdateEnd
End Sub

'=====================================================================
'��ũ�� : ClearSearch
'��� ��Ʈ : Main ��Ʈ
'���� : �˻� ��� ������ �ʱ�ȭ �մϴ�.
'=====================================================================
Public Sub ClearSearch()
    
    Call SetRange '---�� ��� ���� ����
    
    '�˻� ��� ǥ�õǴ� 'DATA' ���� �ʱ�ȭ
    Range("DATA").Clear
    Range("DATA").FormatConditions.Delete '---���Ǻ� ���� ����
    ThisWorkbook.Names("DATA").RefersTo = �˻����
    
    �˻���.ClearContents '--- �˻��� �ʱ�ȭ
    
End Sub

'=====================================================================
'��ũ�� : SearchFile
'��� ��Ʈ : Main ��Ʈ
'���� : ������ �˻��Ͽ� ��ο� ���ϸ��� ��ȸ�մϴ�.
'=====================================================================
Public Sub SearchFile()
    
    Dim varFileNum As Variant
    Dim varFileAdrCheck As Variant
    
    Call SetRange '---�� ��� ���� ����
    
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
        If .SelectedItems.count = 0 Then
        
            MsgBox "������ �������� �ʾҽ��ϴ�."
            Exit Sub
            
        '1�� ���� ���� �� ���� ���ϸ� ����Ʈ ������ �ٿ��ֱ�
        ElseIf .SelectedItems.count = 1 And ���ϸ�.count < 10 And ���ϸ�(1) <> "" Then
            
            varFileNum = InStrRev(.SelectedItems(1), "\") '---'\' �������� ���ϰ�ο� ���ϸ� ����
            ���ϸ�(���ϸ�.count).Offset(1, 0) = Mid(.SelectedItems(1), varFileNum + 1) '---���ϸ� �Է�
            ���ϰ��(���ϰ��.count).Offset(1, 0) = Left(.SelectedItems(1), varFileNum - 1) '---���ϰ�� �Է�
            
            Exit Sub
            
        End If
        
        Union(���ϰ��, ���ϸ�, ��Ʈ��, �Ӹ���).ClearContents '---���� ���� ����Ʈ �ʱ�ȭ
        ��Ʈ��.Validation.Delete '---��Ʈ�� ��Ӵٿ� ����
            
        Call SetRange '---���ϸ� ���� ������
            
        For i = 1 To .SelectedItems.count
            
            '������ ������ ���� ������ �ƴ� ��� ó��
            If InStr(.SelectedItems(i), ".xl") = 0 Then
            
                MsgBox "���� ������ �������ּ���."
                Exit Sub
                
            End If
            
            If i = 11 Then
            
                MsgBox "���õ� ���� ������ 10���� �ʰ��Ͽ� ���� 10���� ���� ����Ʈ�� ȣ��˴ϴ�."
                Exit For
                
            End If
            
            '���ϸ� ����Ʈ�� �� �ٿ��ֱ�
            varFileNum = InStrRev(.SelectedItems(i), "\") '---'\' �������� ���ϰ�ο� ���ϸ� ����
            ���ϸ�(i) = Mid(.SelectedItems(i), varFileNum + 1) '---���ϸ� �Է�
            ���ϰ��(i) = Left(.SelectedItems(1), varFileNum - 1) '---���� ��� �Է�
            
        Next
    End With

End Sub

'=====================================================================
'��ũ�� : SearchSheet
'��� ��Ʈ : Main ��Ʈ
'���� : �Էµ� ���Ͽ��� ��Ʈ���� ��Ӵٿ� �������� ǥ��
'=====================================================================
Public Sub SearchSheet()
    
    'ȣ�� ���� ���� ����
    Dim Obj As Object
    Dim wb As Workbook
    Dim WS As Worksheet
    
    Dim strFile As String '---���� ��� & �̸�
    Dim strSheets() As String '---��Ʈ ����Ʈ ���� �迭
    
    On Error Resume Next
    
    Call UpdateStart
    Call SetRange
        
    '��ο� ���� ���� ��� ó��
    If CheckFile() = True Then
                
        GoTo exit_sub
        
    End If
    
    '�Էµ� ���� ������ŭ �ݺ�
    For i = 1 To ���ϸ�.count
        
        strFile = ���ϰ��(i) & "\" & ���ϸ�(i)
        
        '���� ȣ��
        If CheckFileOpen(strFile) = False Then
                
            Set Obj = GetObject(strFile)
        
        End If
        
        Set wb = Workbooks(Dir(strFile))
        
        Call ObjectList(strFile) '---������Ʈ ����Ʈ ����
        
        ReDim strSheets(1 To wb.Sheets.count) '---�迭 ũ�� ������
        
        For j = 1 To UBound(strSheets)
            
            strSheets(j) = wb.Sheets(j).Name
        
        Next
        
        With ��Ʈ��(i).Validation
            .Delete
            .Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:=Join(strSheets, ",")
            
        End With
        
        ��Ʈ��(i) = strSheets(1)
        
        Erase strSheets '---�迭 �ʱ�ȭ
        
    Next

'���� ó��
exit_sub:

    Call UpdateEnd
    
End Sub


Public Sub ClearFile()

    Call SetRange
    
    Range(���ϰ��, �Ӹ���).ClearContents
    
    ��Ʈ��.Validation.Delete
    
End Sub
