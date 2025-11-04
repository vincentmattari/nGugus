<%
    ' ======================================
    ' 매개변수 생성 헬퍼 함수들
    ' ======================================
    
    ' 간편한 CreateParameter 함수 (Command 객체 없이 사용)
    Function CreateParam(paramName, paramType, paramDirection, paramSize, paramValue)
        Dim param
        Set param = Server.CreateObject("ADODB.Parameter")
        With param
            .Name = paramName
            .Type = paramType
            .Direction = paramDirection
            If paramSize > 0 Then .Size = paramSize
            .Value = paramValue
        End With
        Set CreateParam = param
    End Function
    
    ' 자주 사용하는 매개변수 타입 상수들
    Const DB_BIGINT = 20
    Const DB_INTEGER = 3
    Const DB_VARCHAR = 202
    Const DB_INPUT = 1
    
    ' 문자열 매개변수 생성
    Function CreateStringParam(paramName, paramValue, maxLength)
        Dim param
        Set param = Server.CreateObject("ADODB.Parameter")
        With param
            .Name = paramName
            .Type = adVarWChar
            .Direction = adParamInput
            .Size = maxLength
            .Value = paramValue
        End With
        Set CreateStringParam = param
    End Function
    
    ' 정수 매개변수 생성
    Function CreateIntParam(paramName, paramValue)
        Dim param
        Set param = Server.CreateObject("ADODB.Parameter")
        With param
            .Name = paramName
            .Type = adInteger
            .Direction = adParamInput
            .Value = paramValue
        End With
        Set CreateIntParam = param
    End Function
    
    ' 큰 정수 매개변수 생성 (BIGINT)
    Function CreateBigIntParam(paramName, paramValue)
        Dim param
        Set param = Server.CreateObject("ADODB.Parameter")
        With param
            .Name = paramName
            .Type = adBigInt
            .Direction = adParamInput
            .Value = paramValue
        End With
        Set CreateBigIntParam = param
    End Function
    
    ' 날짜 매개변수 생성
    Function CreateDateParam(paramName, paramValue)
        Dim param
        Set param = Server.CreateObject("ADODB.Parameter")
        With param
            .Name = paramName
            .Type = adDBTimeStamp
            .Direction = adParamInput
            .Value = paramValue
        End With
        Set CreateDateParam = param
    End Function
    
    ' 소수점 매개변수 생성
    Function CreateDecimalParam(paramName, paramValue, precision, scale)
        Dim param
        Set param = Server.CreateObject("ADODB.Parameter")
        With param
            .Name = paramName
            .Type = adNumeric
            .Direction = adParamInput
            .Precision = precision
            .NumericScale = scale
            .Value = paramValue
        End With
        Set CreateDecimalParam = param
    End Function
    
    ' ======================================
    ' 개선된 데이터베이스 함수들
    ' ======================================
    
    ' ======================================
    ' 사용 예시
    ' ======================================
    '
    ' 1. 배열 데이터 조회 (커넥션 풀 사용 - 추천)
    ' Dim params()
    ' ReDim params(1)
    ' Set params(0) = CreateStringParam("@name", "홍길동", 50)
    ' Set params(1) = CreateIntParam("@age", 30)
    ' Dim users = getRsArray("SELECT * FROM users WHERE name = ? AND age = ?", params)
    '
    ' 2. 단일값 조회 (커넥션 풀 사용 - 추천)
    ' Dim countParams()
    ' ReDim countParams(0)
    ' Set countParams(0) = CreateIntParam("@status", 1)
    ' Dim userCount = getRsOne("SELECT COUNT(*) FROM users WHERE status = ?", countParams)
    '
    ' 3. 데이터 수정/삭제/추가 (커넥션 풀 사용 - 추천)
    ' Dim updateParams()
    ' ReDim updateParams(2)
    ' Set updateParams(0) = CreateStringParam("@name", "김철수", 50)
    ' Set updateParams(1) = CreateIntParam("@age", 25)
    ' Set updateParams(2) = CreateBigIntParam("@id", 123)
    ' Dim result = poolSqlCUD("UPDATE users SET name = ?, age = ? WHERE id = ?", updateParams)
    ' If result = 0 Then response.Write("성공") Else response.Write("실패")
    '
    ' 4. 단발성 쿼리 (직접 연결 방식)
    ' Dim quickResult = getRsOneDirect("SELECT @@VERSION", Array())
    '
    ' 5. 기존 코드 호환성 (보안상 권장하지 않음)
    ' Dim oldResult = getRsArrayLegacy("SELECT * FROM users WHERE id = 1")

    ' 개선된 배열 반환 함수 (커넥션 풀 사용)
    Function getRsArray(strSQL, arrParams)
        Dim conn, cmd, rs, arrRs, i
        Set conn = GetDBConnection()
        
        If conn Is Nothing Then
            getRsArray = Array()
            Exit Function
        End If
        
        Set cmd = Server.CreateObject("ADODB.Command")
        
        On Error Resume Next
        With cmd
            .ActiveConnection = conn
            .CommandText = strSQL
            .CommandType = adCmdText
            .Prepared = True
            
            ' 매개변수 추가
            If IsArray(arrParams) Then
                For i = 0 To UBound(arrParams)
                    If IsObject(arrParams(i)) Then
                        .Parameters.Append arrParams(i)
                    End If
                Next
            End If
            
            Set rs = .Execute()
        End With
        
        If Err.Number = 0 And Not rs Is Nothing Then
            If Not rs.EOF Then 
                arrRs = rs.GetRows()
            End If
        End If
        On Error GoTo 0
        
        ' 리소스 정리 (커넥션은 풀에서 관리)
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then rs.Close
            Set rs = Nothing
        End If
        Set cmd = Nothing
        ReleaseDBConnection(conn)
        
        getRsArray = arrRs    
    End Function
    
    ' 기존 방식 (함수 내부에서 매번 연결 생성/해제)
    Function getRsArrayDirect(strSQL, arrParams)
        Dim conn, cmd, rs, arrRs, i
        Set conn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        
        On Error Resume Next
        ' DB 연결
        conn.ConnectionString = "File Name=d:\dbcon\intranet.exc.co.kr.udl;"
        conn.Open
        conn.CursorLocation = adUseClient
        
        If Err.Number = 0 Then
            With cmd
                .ActiveConnection = conn
                .CommandText = strSQL
                .CommandType = adCmdText
                .Prepared = True
                
                ' 매개변수 추가
                If IsArray(arrParams) Then
                    For i = 0 To UBound(arrParams)
                        If IsObject(arrParams(i)) Then
                            .Parameters.Append arrParams(i)
                        End If
                    Next
                End If
                
                Set rs = .Execute()
            End With
            
            If Err.Number = 0 And Not rs Is Nothing Then
                If Not rs.EOF Then 
                    arrRs = rs.GetRows()
                End If
            End If
        End If
        On Error GoTo 0
        
        ' 리소스 정리
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then rs.Close
            Set rs = Nothing
        End If
        Set cmd = Nothing
        If Not conn Is Nothing Then
            If conn.State = adStateOpen Then conn.Close
            Set conn = Nothing
        End If
        
        getRsArrayDirect = arrRs    
    End Function
    
    ' 기존 방식 호환성을 위한 래퍼 함수 (보안상 권장하지 않음)
    Function getRsArrayLegacy(strSQL)
        getRsArrayLegacy = getRsArray(strSQL, Array())
    End Function

    ' 개선된 단일값 반환 함수 (커넥션 풀 사용)
    Function getRsOne(strSQL, arrParams)
        Dim conn, cmd, rs, resultOneValue, i
        Set conn = GetDBConnection()
        
        If conn Is Nothing Then
            getRsOne = Null
            Exit Function
        End If
        
        Set cmd = Server.CreateObject("ADODB.Command")
        
        On Error Resume Next
        With cmd
            .ActiveConnection = conn
            .CommandText = strSQL
            .CommandType = adCmdText
            .Prepared = True
            
            ' 매개변수 추가
            If IsArray(arrParams) Then
                For i = 0 To UBound(arrParams)
                    If IsObject(arrParams(i)) Then
                        .Parameters.Append arrParams(i)
                    End If
                Next
            End If
            
            Set rs = .Execute()
        End With
        
        If Err.Number = 0 And Not rs Is Nothing Then
            If Not rs.EOF Then 
                resultOneValue = rs(0)
            End If
        End If
        On Error GoTo 0
        
        ' 리소스 정리 (커넥션은 풀에서 관리)
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then rs.Close
            Set rs = Nothing
        End If
        Set cmd = Nothing
        ReleaseDBConnection(conn)
        
        getRsOne = resultOneValue
    End Function
    
    ' 기존 방식 (함수 내부에서 매번 연결 생성/해제)
    Function getRsOneDirect(strSQL, arrParams)
        Dim conn, cmd, rs, resultOneValue, i
        Set conn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        
        On Error Resume Next
        ' DB 연결
        conn.ConnectionString = "File Name=d:\dbcon\intranet.exc.co.kr.udl;"
        conn.Open
        conn.CursorLocation = adUseClient
        
        If Err.Number = 0 Then
            With cmd
                .ActiveConnection = conn
                .CommandText = strSQL
                .CommandType = adCmdText
                .Prepared = True
                
                ' 매개변수 추가
                If IsArray(arrParams) Then
                    For i = 0 To UBound(arrParams)
                        If IsObject(arrParams(i)) Then
                            .Parameters.Append arrParams(i)
                        End If
                    Next
                End If
                
                Set rs = .Execute()
            End With
            
            If Err.Number = 0 And Not rs Is Nothing Then
                If Not rs.EOF Then 
                    resultOneValue = rs(0)
                End If
            End If
        End If
        On Error GoTo 0
        
        ' 리소스 정리
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then rs.Close
            Set rs = Nothing
        End If
        Set cmd = Nothing
        If Not conn Is Nothing Then
            If conn.State = adStateOpen Then conn.Close
            Set conn = Nothing
        End If
        
        getRsOneDirect = resultOneValue
    End Function
    
    ' 기존 방식 호환성을 위한 래퍼 함수 (보안상 권장하지 않음)
    Function getRsOneLegacy(strSQL)
        getRsOneLegacy = getRsOne(strSQL, Array())
    End Function

    ' 개선된 CUD 함수 (커넥션 풀 사용)
    Function poolSqlCUD(strSql, arrParams)        
        Dim conn, cmd, resultCUD, i, recordsAffected
        Set conn = GetDBConnection()
        
        If conn Is Nothing Then
            poolSqlCUD = 1  ' 연결 실패
            Exit Function
        End If
        
        Set cmd = Server.CreateObject("ADODB.Command")
        
        On Error Resume Next
        With cmd
            .ActiveConnection = conn
            .CommandText = strSql
            .CommandType = adCmdText
            .Prepared = True
            
            ' 매개변수 추가
            If IsArray(arrParams) Then
                For i = 0 To UBound(arrParams)
                    If IsObject(arrParams(i)) Then
                        .Parameters.Append arrParams(i)
                    End If
                Next
            End If
            
            .Execute recordsAffected
        End With
        
        If Err.Number <> 0 Then
            resultCUD = 1   '에러
            Err.Clear
        Else
            resultCUD = 0   '정상
        End If
        On Error GoTo 0
        
        ' 리소스 정리 (커넥션은 풀에서 관리)
        Set cmd = Nothing
        ReleaseDBConnection(conn)
        
        poolSqlCUD = resultCUD
    End Function
    
    ' 기존 방식 (함수 내부에서 매번 연결 생성/해제)
    Function poolSqlCUDDirect(strSql, arrParams)        
        Dim conn, cmd, resultCUD, i, recordsAffected
        Set conn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        
        On Error Resume Next
        ' DB 연결
        conn.ConnectionString = "File Name=d:\dbcon\intranet.exc.co.kr.udl;"
        conn.Open
        conn.CursorLocation = adUseClient
        
        If Err.Number = 0 Then
            With cmd
                .ActiveConnection = conn
                .CommandText = strSql
                .CommandType = adCmdText
                .Prepared = True
                
                ' 매개변수 추가
                If IsArray(arrParams) Then
                    For i = 0 To UBound(arrParams)
                        If IsObject(arrParams(i)) Then
                            .Parameters.Append arrParams(i)
                        End If
                    Next
                End If
                
                .Execute recordsAffected
            End With
            
            If Err.Number <> 0 Then
                resultCUD = 1   '에러
                Err.Clear
            Else
                resultCUD = 0   '정상
            End If
        Else
            resultCUD = 1   '연결 에러
            Err.Clear
        End If
        On Error GoTo 0
        
        ' 리소스 정리
        Set cmd = Nothing
        If Not conn Is Nothing Then
            If conn.State = adStateOpen Then conn.Close
            Set conn = Nothing
        End If
        
        poolSqlCUDDirect = resultCUD
    End Function
    
    ' 기존 방식 호환성을 위한 래퍼 함수 (보안상 권장하지 않음)
    Function poolSqlCUDLegacy(strSql)
        poolSqlCUDLegacy = poolSqlCUD(strSql, Array())
    End Function
    
%>