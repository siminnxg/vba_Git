Attribute VB_Name = "ModuleCheat1"
Option Explicit
'###################################################
'Cheat1 RequestCreateItem ġƮŰ ���� ���
'###################################################

'======================================================================================================
'ġƮŰ1 [Cheat ����] ��ư Ŭ�� �� ����


Public Sub Cheat1()
    
    '# ���� ���� #
    Call UpdateStart
    Call SetRange
    
    '���õ� Ű�� ���� �� ����
    If �˻����_����.Value = "" Then
    
        MsgBox "���õ� Key�� �����ϴ�."
        
        GoTo ����
        
    End If
    
    '���� ������ ��ȸ
    Call SQLFileLoad(�˻����, Ÿ��.ListColumns("����").DataBodyRange)
    
    'ġƮŰ ����
    Call CheatCreatItem
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� �Է�] ��ư�� Ŭ�����ּ���." '---��ܿ� �ȳ� ���� ǥ��
    
����:

    Call UpdateEnd
    
End Sub

'======================================================================================================
'SQL�� ���� ������ ��ȸ


Public Function SQLFileLoad(cell As Range, rngFileName As Range)
    
    '# ���� ���� #
    
    Dim objDB As Object 'ADODB ��ü ������ ����
    Dim obj As Object '������ ��ü ���� ����
    Dim strSQL As String 'SQL�� ���� ����
    Dim strFilePath As String '���� ���
    
    Dim strWhere As String 'Where ���� ����
    Dim strRuneData As Variant 'RuneUIData �������� ã�� �� ���� �迭
    Dim strFileName As String '��ȸ�� ���� �̸� ���� ����
    Dim strFolder As String '���� ��� ���� ����
    
    Dim rngFindCell As Range '�˻��� Key�� ��ġ ����
    Dim rngRuneCell As Range 'RuneData ��Ʈ���� key�� ��ġ ����
    
    '# ���� ���� #
    
    '���� �߻� �� ���� �� ����
    On Error Resume Next
    
    strFolder = LatestFolder '---�ֽ� ���� ��� ����
    
    '���� ��ΰ� ���� �� ����
    If strFolder = "" Then
    
        Exit Function
        
    End If
    
    '���õ� Key ������ ���� Where �������� ��ȯ
    For i = 1 To cell.Cells.Count
    
        strWhere = strWhere & "'" & cell(i).Value & "',"
        
    Next
    
    '���� ������ŭ �ݺ�
    For i = 1 To rngFileName.Cells.Count
        
        strFileName = rngFileName(i).Value '---���� �̸� ����
        
        strFilePath = strFolder & "\" & strFileName & ".xlsx" '---���� ��� ����
         
        '�Էµ� ������ �������� ���� �� ����
        If CheckFile(strFilePath) = True Then
        
            Exit Function
            
        End If
        
        'SQL ���� �ۼ�
        '������ �� ������ �� �� ����
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
            
            ReDim strRuneData(cell.Cells.Count - 1, 1) '---�˻��� ������ŭ �迭 ũ�� ����
            
            '��ȸ�� ���� ���� �� ����
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
                
                rngFindCell.Offset(0, 1) = rngRuneCell.Offset(0, 1).Value '---
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

'======================================================================================================
'ġƮŰ ����


Public Sub CheatCreatItem()
    
    '# ���� ���� #
    
    Dim InItemType As Variant '---������ Ÿ�� ���� ����
    Dim InTemplateId As Variant '---������ TID ���� ����
    Dim InCount As Variant '---������ ���� ���� ����
    Dim InLevel As Variant '---������ ���� ���� ����
        
        
    '# ���� ���� #
    
    '���� ���� ȣ��
    Call SetRange
    
    ġƮŰ.ClearContents '---ġƮŰ ���� �ʱ�ȭ
    
    '���õ� KEY ������ŭ �ݺ�
    For i = 0 To �˻����.Cells.Count - 1
        
        With �˻����(i + 1)
            
            InTemplateId = .Offset(0, 1).Value '---TID �� ����
            
            '������ ���� ������ Ÿ�� ����
            '����, �����ǰ, ������
            If .Offset(0, 2).Value = "RangedWeaponData" Or .Offset(0, 2).Value = "AccessoryData" Or .Offset(0, 2).Value = "ReactorData" Then
            
                InItemType = 2
            
            '���
            ElseIf .Offset(0, 2).Value = "ConsumableItemData" Then
            
                InItemType = 3
            
            '��
            ElseIf .Offset(0, 2).Value = "RuneUIData" Then
            
                InItemType = 4
             
            'Ŀ���͸���¡
            ElseIf .Offset(0, 2).Value = "CustomizingItemData" Then
            
                InItemType = 7
                
            End If
            
            InCount = .Offset(0, 3).Value '---������ ���� ����
            
            '���� �� 1�� �Է�
            If InCount = 0 Then
            
                InCount = 1
                
            End If
                        
            InLevel = .Offset(0, 4).Value '---������ ���� ����
            
            '���� �� ���� 100 �Է�
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
    
    '���� ������ ����Ʈ ǥ��
    Call LoadTxt
    
End Sub
