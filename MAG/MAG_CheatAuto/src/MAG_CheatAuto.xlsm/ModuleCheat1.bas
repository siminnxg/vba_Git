Attribute VB_Name = "ModuleCheat1"
Option Explicit
'###################################################
'Cheat1 RequestCreateItem ġƮŰ ���� ���
'###################################################

'===================================================
'������ ���� [Cheat ����] ��ư
Public Sub Cheat1()
    
    '# ���� ����
    Call UpdateStart
    Call SetRange
    
    '���õ� Ű�� ������ ����
    If �˻����_����.Value = "" Then
        MsgBox "���õ� Key�� �����ϴ�."
        GoTo Exit_Sub
    End If
    
    '���� ������ ��ȸ
    Call SQLFileLoad(�˻����, Ÿ��.ListColumns("����").DataBodyRange)
    
    'ġƮŰ ����
    Call CheatCreatItem
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� ����] ��ư�� Ŭ�����ּ���."
    
Exit_Sub:
    Call UpdateEnd
    
End Sub

'===================================================
'SQL�� ���� ������ ��ȸ
Public Function SQLFileLoad(cell As Range, rngFileName As Range)
    
    '# ���� ����
    Dim objDB As Object 'ADODB ��ü ������ ����
    Dim obj As Object '������ ��ü ���� ����
    Dim strSQL As String 'SQL�� ���� ����
    Dim strFilePath As String '���� ���
    
    Dim strWhere As String '---Where ���� ����
    Dim rngFindCell As Range '---�˻��� Key�� ��ġ ����
    Dim rngRuneCell As Range '---RuneData ��Ʈ���� key�� ��ġ ����
    Dim strRuneData As Variant '---RuneUIData �������� ã�� �� ���� �迭
    Dim strFileName As String
    Dim strFolder As String
    
    '# ���� ����
    On Error Resume Next
        
    strFolder = LatestFolder
    
    If strFolder = "" Then
        Exit Function
        
    End If
    
    '���õ� Key ������ ���� Where �������� ��ȯ
    For i = 1 To cell.Cells.Count
        strWhere = strWhere & "'" & cell(i).Value & "',"
    Next
    
    '���� ������ŭ �ݺ�
    For i = 1 To rngFileName.Cells.Count
        
        strFileName = rngFileName(i).Value
        
        strFilePath = strFolder & "\" & strFileName & ".xlsx" '---���� ��� ����
         
        If CheckFile(strFilePath) = True Then
            Exit Function
        End If
        
        'SQL ���� �ۼ�
        '�� �����ʹ� ���� �ۼ�
        If rngFileName(i) = "RuneUIData" Then
            
            j = 0
            
            'RuneUIData �������� ������ ��ȸ
            strSQL = " SELECT * " & _
                 " FROM [Data$] " & _
                 " WHERE TitleStringKey IN (" & strWhere & ")"
                 
            'OLEDB ����
            Set objDB = CreateObject("ADODB.Connection")
            objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & strFilePath & ";" & _
                      "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
            
            Set obj = CreateObject("ADODB.Recordset")
            obj.Open strSQL, objDB
            
            ReDim strRuneData(cell.Cells.Count - 1, 1)
            
            '��ȸ�� ���� �ִ� ��� ��Ʈ�� ǥ��
            Do Until obj.EOF
            
                If j = 0 Then
                    strRuneData(j, 0) = obj("TitleStringKey")
                    strRuneData(j, 1) = obj.Fields(0)
                    
                    j = j + 1
                    
                '������ Ű�� ���� ������ �����ؼ� �ߺ� ����
                ElseIf strRuneData(j - 1, 0) <> obj("TitleStringKey") Then
                    
                    strRuneData(j, 0) = obj("TitleStringKey")
                    strRuneData(j, 1) = obj.Fields(0)
                
                    j = j + 1
                End If

                obj.MoveNext
            Loop
            
            '��ü ���� ����
            obj.Close
            objDB.Close
            Set obj = Nothing
            Set objDB = Nothing
            
            For k = 0 To j - 1
                Set rngFindCell = cell.Find(strRuneData(k, 0), Lookat:=xlWhole)
                Set rngRuneCell = Sheets("RuneData").UsedRange.Find(strRuneData(k, 1), Lookat:=xlWhole)
                
                rngFindCell.Offset(0, 1) = rngRuneCell.Offset(0, 1).Value
                rngFindCell.Offset(0, 2) = rngFileName(i) '---��ȸ�� ���ϸ� �Է�
            Next
                  
        Else
            'DATA ��Ʈ���� ���ǿ� �´� TemplateId, StringId ���� ������ ����
            strSQL = " SELECT TemplateId, StringId " & _
                 " FROM [DATA$] " & _
                 " WHERE StringId IN (" & strWhere & ")"
            
            'OLEDB ����
            Set objDB = CreateObject("ADODB.Connection")
            objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & strFilePath & ";" & _
                      "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
            
            Set obj = CreateObject("ADODB.Recordset")
            obj.Open strSQL, objDB
            
            '��ȸ�� ���� �ִ� ��� ��Ʈ�� ǥ��
            Do Until obj.EOF
                
                Set rngFindCell = cell.Find(obj("StringId"), Lookat:=xlWhole)
                rngFindCell.Offset(0, 1) = obj("TemplateId") '---TID�� �Է�
                rngFindCell.Offset(0, 2) = rngFileName(i) '---��ȸ�� ���ϸ� �Է�
                
                obj.MoveNext
            Loop
            
            '��ü ���� ����
            obj.Close
            objDB.Close
            Set obj = Nothing
            Set objDB = Nothing
            
        End If
    Next
    
End Function

'===================================================
'[Cheat ����] ��ư
Public Sub CheatCreatItem()
    
    Dim InItemType As Variant
    Dim InTemplateId As Variant
    Dim InCount As Variant
    Dim InLevel As Variant
    
    Call SetRange
    
    ġƮŰ.ClearContents
    
    For i = 0 To �˻����.Cells.Count - 1
        
        With �˻����(i + 1)
            
            InTemplateId = .Offset(0, 1).Value
            
            '������ ���� ������ Ÿ�� ����
            If .Offset(0, 2).Value = "RangedWeaponData" Or .Offset(0, 2).Value = "AccessoryData" Or .Offset(0, 2).Value = "ReactorData" Then
                InItemType = 2
                
            ElseIf .Offset(0, 2).Value = "ConsumableItemData" Then
                InItemType = 3
                
            ElseIf .Offset(0, 2).Value = "RuneUIData" Then
                InItemType = 4
                
            ElseIf .Offset(0, 2).Value = "CustomizingItemData" Then
                InItemType = 7
            End If
            
            '������ ���� ���� (���� �� 1)
            InCount = .Offset(0, 3).Value
            If InCount = 0 Then
                InCount = 1
            End If
            
            '������ ���� ���� (���� �� 100)
            InLevel = .Offset(0, 4).Value
            If InLevel = 0 Then
                InLevel = 100
            End If
        
        End With
        
        '������ ID ���� �� �ȳ� ���� ǥ��
        If InTemplateId = 0 Then
            ġƮŰ_����.Offset(i, 0).Value = "��ȸ�� TID�� �������� �ʽ��ϴ�."
        
        'ġƮŰ �Է�
        Else
            ġƮŰ_����.Offset(i, 0).Value = "RequestCreateItem " & InItemType & " " & InTemplateId & " " & _
                                        InCount & " " & InLevel
        End If
    
    Next
    
End Sub
