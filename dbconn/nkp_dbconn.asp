<%
'---------------------------------------------------------------------------
'클래스 명 : DBConn
' Description : DB Connection 클래스화
' Example :
    'Set oDB = new DBConn

    'getQuery = "" &_ "select * " &_ "from ps_member " &_ "order by " &_ " mem_id"

    'Set rs = oDB.ExecuteQuery(query)

    'Do While Not rs.EOF
    '   ...
    'Loop

    'oDB.Close
' Output :
'---------------------------------------------------------------------------
Class DBConn
    Private objConn 'as ADODB.Connection
    Public connString

    Dim DBHost, DBName, DBUser, DBPass

    '클래스 생성자
    Sub Class_Initialize()
        Set objConn = Nothing

        DBHost = "localhost"
        DBName = "nkp"
		'DBName = "nkp_test"	'개발 테스트 DB
        DBUser = "nkp"
        DBPass = "nkp2014"

        'connString = "DRIVER={MySQL ODBC 5.3 ansi Driver};Password=" & DBPass &_
        '    ";Persist Security Info=False;User ID=" & DBUser &_
        '    ";Initial Catalog=" & DBName & ";Data Source=" & DBHost
        connString = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER="&DBHost&_
            ";DATABASE="&DBName&";UID="&DBUser&";PWD="&DBPass&"; "
    End Sub

    '클래스 소멸자
    Sub Class_Terminate()
        Set objConn = Nothing
    End Sub

    'Connection 프로퍼티 Get
    Public Property Get Connection()
        'as ADODB.Connection
        If objConn Is Nothing Then
            Set objConn = Server.CreateObject("ADODB.Connection")
            objConn.Open connString
        End If

        Set Connection = objConn
    End Property

    '레코드셋 질의
    Public Function ExecuteQuery(strSQL)
        Dim objRs 'as ADODB.Recordset

        Set objRs = Server.CreateObject("ADODB.Recordset")

        On Error Resume Next

        objRs.CursorLocation = 3
        objRs.Open strSQL, Me.Connection, 0

        If Err.Number <> 0 Then
            Response.Write "<b>데이터베이스 에러</b> (ExecuteQuery)<br>" &_
                "질의어 : " & strSQL &_
                Err.Description

            objRs.Close

            Set objRs = Nothing
            Me.Close

            Response.End
        End If

        On Error GoTo 0
        Set ExecuteQuery = objRs
    End Function

    '업데이트 질의
    Public Sub ExecuteCommand(strSQL)
        On Error Resume Next

        Me.Connection.Execute strSQL

        If Err.Number <> 0 Then
            Response.Write "<b>데이터베이스 에러</b> (ExecuteCommand)<br>" &_
                "질의어 : " & strSQL &_
                Err.Description

            Response.End
            Me.Close
        End If

        On Error GoTo 0
    End Sub

    '데이터베이스 연결 닫기
    Public Sub Close()
        Set objConn = Nothing
    End Sub
End Class
%>
