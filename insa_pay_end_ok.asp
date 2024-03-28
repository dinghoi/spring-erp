<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	pmg_yymm = request.form("pmg_yymm1")
	etc_code = request.form("etc_code")

	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	sql = "Update emp_etc_code set emp_payend_date='"&pmg_yymm&"',emp_payend_yn='Y' where emp_etc_code = '"&etc_code&"'"
	dbconn.execute(sql)
	
	end_msg = pmg_yymm + " 월 마감등록 되었습니다...."
	
	response.write"<script language=javascript>"
	'response.write"alert('마감등록 되었습니다....');"	
	response.write"alert('"&end_msg&"');"	
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	
	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>
	