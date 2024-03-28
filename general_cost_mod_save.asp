<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	slip_seq = request.form("slip_seq")
	slip_date = request.form("slip_date")
	accountitem = request.form("account")
	i=instr(1,accountitem,"-")
	account = mid(accountitem,1,i-1)
	account_item = mid(accountitem,i+1)
	price = int(request.form("price"))
	customer = request.form("customer")
	pay_yn = request.form("pay_yn")
	cancel_yn = request.form("cancel_yn")
	confirm_yn = request.form("confirm_yn")
  	cost_vat = 0
	cost = price
	if pay_yn = "Y" then
		cancel_yn = "N"
	end if

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "update general_cost set account='"&account&"',account_item='"&account_item&"',price="&price&",cost="&cost&",cost_vat="&cost_vat&",customer='"&customer&"',cancel_yn='"&cancel_yn&"',confirm_yn='"&confirm_yn&"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where slip_date='"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	dbconn.execute(sql)	  

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
