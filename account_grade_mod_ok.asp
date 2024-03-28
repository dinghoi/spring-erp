<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<%
	user_id = request.form("user_id")
	account_grade = request.form("account_grade")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	sql = "Update memb set account_grade='"&account_grade&"' where user_id = '"&user_id&"'"
	dbconn.execute(sql)

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
