<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	approve_no = request.form("approve_no")
	cust_no1 = request.form("cust_no1")
	cust_no2 = request.form("cust_no2")
	cust_no3 = request.form("cust_no3")
	customer_no = cstr(cust_no1) + "-" + cstr(cust_no2) + "-" + cstr(cust_no3)
	customer = request.form("customer")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	sql = "Update card_slip set customer='"&customer&"',customer_no='"&customer_no&"',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where approve_no = '"&approve_no&"'"
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
