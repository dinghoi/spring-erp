<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	holiday = request.form("holiday")
	holiday_memo = request.form("holiday_memo")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	sql = "insert into holiday (holiday,holiday_memo,reg_id,reg_date) values ('"&holiday&"','"&holiday_memo&"','"&user_id&"',now())"
	dbconn.execute(sql)
		
	response.write"<script language=javascript>"
	response.write"alert('입력 완료 되었습니다....');"		
	response.Redirect "holi_mg.asp"
	response.write"</script>"	
	Response.End
	dbconn.Close()
	Set dbconn = Nothing
%>
