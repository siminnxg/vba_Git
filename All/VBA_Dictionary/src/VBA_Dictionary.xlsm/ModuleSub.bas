Attribute VB_Name = "ModuleSub"
Option Explicit

Sub 연결제거()

    Dim conn As Object
    Dim connName As String
    
    '---모든 연결을 순회
    For Each conn In ActiveWorkbook.Connections
        connName = conn.Name '---연결 이름 가져오기
        
        '---연결로 시작하는 이름의 연결 제거
        If connName Like "연결*" Then
        
            conn.Delete
            
        End If
    Next conn
    
End Sub
