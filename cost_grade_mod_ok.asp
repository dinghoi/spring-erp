<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<%
	user_id = request.form("user_id")
	cost_grade = request.form("cost_grade")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	sql = "Update memb set cost_grade='"&cost_grade&"' where user_id = '"&user_id&"'"
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
