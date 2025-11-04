<!--METADATA TYPE="typelib" NAME="ADODB Type Library" File="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<%     
    ' ======================================
    ' DB 커넥션 관리 클래스
    ' ======================================
    
    ' 글로벌 커넥션 풀 변수
    Dim g_dbConnPool
    Set g_dbConnPool = Nothing  ' 명시적으로 Nothing으로 초기화
    
    ' 커넥션 풀 초기화
    Sub InitializeDBPool()
        On Error Resume Next
        If g_dbConnPool Is Nothing Or IsEmpty(g_dbConnPool) Then
            Set g_dbConnPool = Server.CreateObject("Scripting.Dictionary")
            If Err.Number <> 0 Then
                Response.Write "DB 풀 초기화 오류: " & Err.Description & "<br>"
                Err.Clear
            End If
        End If
        On Error GoTo 0
    End Sub
    
    ' 커넥션 가져오기 (풀에서 재사용 또는 새로 생성)
    Function GetDBConnection()
        On Error Resume Next
        InitializeDBPool()
        
        If Err.Number <> 0 Then
            Response.Write "InitializeDBPool 오류: " & Err.Description & "<br>"
            Set GetDBConnection = Nothing
            Exit Function
        End If
        
        Dim connKey, conn
        connKey = "main_connection_" & Session.SessionID
        
        ' 기존 연결 확인
        If Not g_dbConnPool Is Nothing Then
            If g_dbConnPool.Exists(connKey) Then
                Set conn = g_dbConnPool(connKey)
                ' 연결 상태 확인
                If Not conn Is Nothing And conn.State = 1 Then  ' adStateOpen = 1
                    Set GetDBConnection = conn
                    On Error GoTo 0
                    Exit Function
                Else
                    ' 끊어진 연결 제거
                    g_dbConnPool.Remove(connKey)
                End If
            End If
        End If
        
        ' 새 연결 생성
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.ConnectionString = "File Name=d:\dbcon\intranet.exc.co.kr.udl;"
        conn.Open
        conn.CursorLocation = 3  ' adUseClient = 3
        
        If Err.Number = 0 Then
            If Not g_dbConnPool Is Nothing Then
                g_dbConnPool.Add connKey, conn
            End If
            Set GetDBConnection = conn
        Else
            Response.Write "DB 연결 오류: " & Err.Description & "<br>"
            Set GetDBConnection = Nothing
            Err.Clear
        End If
        On Error GoTo 0
    End Function
    
    ' 커넥션 해제 (실제로는 풀에 보관)
    Sub ReleaseDBConnection(conn)
        ' 연결을 닫지 않고 풀에 보관 (세션 종료시 자동 정리)
    End Sub
    
    ' 커넥션 풀 정리
    Sub CleanupDBPool()
        If Not g_dbConnPool Is Nothing Then
            Dim key
            For Each key In g_dbConnPool.Keys
                If Not g_dbConnPool(key) Is Nothing Then
                    If g_dbConnPool(key).State = adStateOpen Then
                        g_dbConnPool(key).Close
                    End If
                    Set g_dbConnPool(key) = Nothing
                End If
            Next
            g_dbConnPool.RemoveAll
        End If
    End Sub

%>