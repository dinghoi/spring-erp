<%
'---------------------------------------------------------------------------
'Ŭ���� �� : DBConn
' Description : DB Connection Ŭ����ȭ
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

    'Ŭ���� ������
    Sub Class_Initialize()
        Set objConn = Nothing

        DBHost = "localhost"
        DBName = "nkp"
		'DBName = "nkp_test"	'���� �׽�Ʈ DB
        DBUser = "nkp"
        DBPass = "nkp2014"

        'connString = "DRIVER={MySQL ODBC 5.3 ansi Driver};Password=" & DBPass &_
        '    ";Persist Security Info=False;User ID=" & DBUser &_
        '    ";Initial Catalog=" & DBName & ";Data Source=" & DBHost
        connString = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER="&DBHost&_
            ";DATABASE="&DBName&";UID="&DBUser&";PWD="&DBPass&"; "
    End Sub

    'Ŭ���� �Ҹ���
    Sub Class_Terminate()
        Set objConn = Nothing
    End Sub

    'Connection ������Ƽ Get
    Public Property Get Connection()
        'as ADODB.Connection
        If objConn Is Nothing Then
            Set objConn = Server.CreateObject("ADODB.Connection")
            objConn.Open connString
        End If

        Set Connection = objConn
    End Property

    '���ڵ�� ����
    Public Function ExecuteQuery(strSQL)
        Dim objRs 'as ADODB.Recordset

        Set objRs = Server.CreateObject("ADODB.Recordset")

        On Error Resume Next

        objRs.CursorLocation = 3
        objRs.Open strSQL, Me.Connection, 0

        If Err.Number <> 0 Then
            Response.Write "<b>�����ͺ��̽� ����</b> (ExecuteQuery)<br>" &_
                "���Ǿ� : " & strSQL &_
                Err.Description

            objRs.Close

            Set objRs = Nothing
            Me.Close

            Response.End
        End If

        On Error GoTo 0
        Set ExecuteQuery = objRs
    End Function

    '������Ʈ ����
    Public Sub ExecuteCommand(strSQL)
        On Error Resume Next

        Me.Connection.Execute strSQL

        If Err.Number <> 0 Then
            Response.Write "<b>�����ͺ��̽� ����</b> (ExecuteCommand)<br>" &_
                "���Ǿ� : " & strSQL &_
                Err.Description

            Response.End
            Me.Close
        End If

        On Error GoTo 0
    End Sub

    '�����ͺ��̽� ���� �ݱ�
    Public Sub Close()
        Set objConn = Nothing
    End Sub
End Class
%>
