Attribute VB_Name = "ModuleMain"
Option Explicit

Sub Student_Info()
    
    Dim �л�����_���� As Range
    Dim �л����_���� As Range
    
    Dim cell1 As Range
    Dim cell2 As Range
    Dim rngFun As Range
    Dim cell_now As Range
    
    Dim cnt_header As Variant
    Dim cnt_st As Variant
    Dim stNum As Variant
    Dim column As String
    Dim i As Variant
    
    Dim ws As Worksheet
    
    Set ws = Sheets("������ ó��")
    
    Set �л�����_���� = ws.Range("AF5")
    Set �л����_���� = ws.Range("AB6")
    
    
    cnt_st = 0
    cnt_header = 0
    stNum = 1
    
    Set cell1 = �л�����_����
    
    Do While cell1 <> ""
        
        �л����_����.Offset(cnt_header, 0) = "�л�" & stNum
                
        Set cell2 = cell1
        
        Do While cell2.Value <> "��END��"
            
            If InStr(cell2, "��") Then
                
                '������ ���� ���� ����
                Set rngFun = �л����_����.Offset(cnt_header, 1)
                
                �л����_����.Offset(cnt_header, 1) = cell2
                
                '�� ���ĺ� ����
                column = Split(cell1.Offset(0, 3).Address(False, False), cell1.Offset(0, 3).Row)(0)
                
                '3�� �Լ� �Է�
                �л����_����.Offset(cnt_header, 2).Formula = "=LET(" & _
                                                                "titleCell, MATCH(" & rngFun.Address & ", " & column & ":" & column & ", 0)," & _
                                                                "nextTitleCell, IFERROR(AGGREGATE(15, 6, ROW($" & column & "$2:$" & column & "$100)/(LEFT($" & column & "$2:$" & column & "$100,1)=""��"")/(ROW($" & column & "$2:$" & column & "$100)>titleCell), 1), ROWS(" & column & ":" & column & ")+1)," & _
                                                                "startRow, titleCell + 1," & _
                                                                "endRow, nextTitleCell - 1," & _
                                                                "colStart, COLUMN($" & column & "$2)," & _
                                                                "colEnd, colStart + 1," & _
                                                                "ADDRESS(startRow, colStart) & "":"" & ADDRESS(endRow, colEnd)" & _
                                                                ")"
                                
                
                '������ ���� ��ȸ�� ���� ����
                cnt_header = cnt_header + 1
                
            End If
            
            Set cell2 = cell2.Offset(1, 0)
        Loop
                
        stNum = stNum + 1 '�л� ��
        cnt_st = cnt_st + 5 '�л� ���� ���� ��ġ
        
        Set cell1 = �л�����_����.Offset(0, cnt_st)
        
    Loop
    
End Sub

