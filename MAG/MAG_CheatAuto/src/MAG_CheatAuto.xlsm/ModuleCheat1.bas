Attribute VB_Name = "ModuleCheat1"
Option Explicit
'###################################################
'������ ���� ġƮŰ ���
'###################################################

Sub FindTID()
    
    '# ���� ����
    Dim rngFileName As Range '���ϰ�� & ���ϸ�
    Dim rngFindType As Range
    
    '# ���� ����
    Call UpdateStart
    Call SetRange
    
    If �˻����_����.Value = "" Then
        MsgBox "���õ� Key�� �����ϴ�."
        GoTo Exit_Sub
    End If
        
    Call SQLFileLoad(�˻����, Ÿ��.ListColumns("����").DataBodyRange)
    
    Call CheatCreatItem
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� ����] ��ư�� Ŭ�����ּ���."
    
Exit_Sub:
    Call UpdateEnd
    
End Sub

Public Sub SelectKey()
    
    Call SetRange
    
    'Ű ����� ��ȸ�ϸ� ���õ� �� Ȯ��
    For Each cell In Ű���
        If cell.Offset(0, -1).Interior.Color = vbRed Then
            
            '�˻� ��� ����ִ� ���
            If �˻����_����.Value = "" Then
                �˻����_����.Value = cell.Value
                
            '�˻� ��� ���� �����ϴ� ���
            Else
                Set �˻����_�� = �˻����_��.Offset(1, 0)
                �˻����_��.Value = cell.Value
            End If
        End If
    Next
    
    Ű���.Offset(0, -1).Interior.Color = vbWhite '---�̵� �Ϸ� �� KEY ���� �ʱ�ȭ

End Sub

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
    
    '# ���� ����
    'On Error Resume Next
    
    '���õ� Key ������ ���� Where �������� ��ȯ
    For i = 1 To cell.Cells.Count
        strWhere = strWhere & "'" & cell(i).Value & "',"
    Next
    
    '���� ������ŭ �ݺ�
    For i = 1 To rngFileName.Cells.Count
        
        strFileName = rngFileName(i).Value
        
        strFilePath = ���ϰ��.Value & "\" & strFileName & ".xlsx" '---���� ��� ����
        
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

'RequestCreateItem InItemType, InTemplateld, InCount, InLevel
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
            ġƮŰ_����.Offset(i, 0).Value = "��ȸ�� TID�� �������� �ʽ��ϴ�."
        Else
        
            ġƮŰ_����.Offset(i, 0).Value = "RequestCreateItem " & InItemType & " " & InTemplateId & " " & _
                                        InCount & " " & InLevel
        End If
    
    Next
    
End Sub

Public Sub ClearKEYList()

    Call SetRange
    
    �˻����.Resize(, 5).ClearContents
    
End Sub

Public Sub ClearCheatList()
    
    Call SetRange
    
    ġƮŰ.ClearContents
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� ����] ��ư�� Ŭ�����ּ���."
    
End Sub

Public Sub ClearSearchList()
    
    Call SetRange
    
    Ű���.Offset(0, -1).Interior.Color = vbWhite
    
End Sub

Public Sub WriteCheat()
    
    Dim path As String
    
    Call SetRange
    
    If IsEmpty(ġƮŰ_����) Then
        MsgBox "������ ġƮŰ�� �����ϴ�."
        Exit Sub
    End If
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    Open path For Output As #1
        
        Print #1, "<Mag_CreatItem>"
        
        'ġƮŰ ������ ���鼭 �ݺ�
        For Each cell In ġƮŰ
            
            '��ȸ�� TID~~ �� ����
            If InStr(cell.Value, "��ȸ��") = 0 Then
                Print #1, cell.Value '---�ۼ��� ġƮ �޸��忡 ���
            End If
            
        Next
    Close
    
    ġƮŰ_����.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " <Mag_CreatItem>"

End Sub
