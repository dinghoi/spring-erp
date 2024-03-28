<%@LANGUAGE="VBSCRIPT"%>
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"
%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
    On Error Resume Next

    Dim result : result = "fail"
    Dim sql
    Dim orgColumn, sqlWhere

    user_id = request("user_id")
    grade = request("grade")

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")

	dbconn.open DbConnect

	sql = " UPDATE memb                         "&chr(13)&_
	      "    SET grade = '" & grade & "'      "&chr(13)&_
	      "  WHERE user_id = '" & user_id & "'  "
    'Response.write sql&chr(13)
    dbconn.execute(sql)
    result = "succ"

	Dbconn.close : Set Dbconn = Nothing

    If Err.number<>0 Then
        result = "error"
    End IF

    If Trim(result&"")<>"" Then
        result = "{""result"" : """ & result & """}"
    End If

    Response.write result
%>