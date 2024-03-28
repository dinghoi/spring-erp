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

    card_no = request("card_no")

	Set Dbconn = Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")	
	
	dbconn.open DbConnect

	sql = " SELECT count(*) cnt                 "&chr(13)&_
	      "   FROM card_owner                   "&chr(13)&_
	      "  WHERE card_no = '" & card_no & "'  "&chr(13)
    'Response.Write  "<pre>" & Sql &"</pre>"
    Set Rs = Dbconn.Execute (sql)
	
    total_record = cint(Rs(0)) 'Result.RecordCount
    
    result = "succ"

	Dbconn.close : Set Dbconn = Nothing

    If Err.number<>0 Then
        result = "error"
    End IF

    If Trim(result&"")<>"" Then
        result = "{""result"" : """ & result & """, ""total_record"" : """ & total_record & """}"
    End If

    Response.write result
%>