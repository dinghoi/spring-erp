<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	u_type = request.form("u_type")
	view_c = request.form("view_c")
	field_view = request.form("field_view")
	card_upjong = request.form("card_upjong")
	account_view = request.form("account_view")
	i=instr(1,account_view,"-")
	account = mid(account_view,1,i-1)
	account_item = mid(account_view,i+1)
	tax_yn = request.form("tax_yn")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	if	u_type = "U" then
		sql = "Update card_upjong set account='"&account&"',account_item='"&account_item&"',tax_yn='"&tax_yn&"',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where card_upjong = '"&card_upjong&"'"
		dbconn.execute(sql)
	  else
		sql="insert into card_upjong (card_upjong,account,account_item,tax_yn,reg_id,reg_name,reg_date) values ('"&card_upjong&"','"&account&"','"&account_item&"','"&tax_yn&"','"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)
	end if	

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.Redirect "card_cust_upjong_mg.asp?view_c="&view_c&"&field_view="&field_view&"&ck_sw=y"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
