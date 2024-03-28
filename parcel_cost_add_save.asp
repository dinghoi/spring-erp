<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	slip_seq = request.form("slip_seq")
	slip_date = request.form("slip_date")
	old_date = request.form("old_date")
	slip_gubun = "운반비"
	company = request.form("company")
	emp_company = request.form("emp_company")
	bonbu = request.form("bonbu")
	saupbu = request.form("saupbu")
	team = request.form("team")
	org_name = request.form("org_name")
	account = request.form("account")
	account_item = account
	pay_method = "현금"
	price = int(request.form("price"))
	cost = int(request.form("cost"))
	cost_vat = int(request.form("cost_vat"))
	vat_yn = "Y"
	customer = request.form("customer")
	customer_no = request.form("customer_no")
	pay_yn = "N"
	slip_memo = request.form("slip_memo")
	end_yn = request.form("end_yn")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "delete from general_cost where slip_date ='"&old_date&"' and slip_seq='"&slip_seq&"'"
		dbconn.execute(sql)
	end if

	sql="select max(slip_seq) as max_seq from general_cost where slip_date='"&slip_date&"'"
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_seq"))  then
		slip_seq = "001"
	  else
		max_seq = "00" + cstr((int(rs("max_seq")) + 1))
		slip_seq = right(max_seq,3)
	end if

	sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,emp_company,bonbu,saupbu,team,org_name,company,account,account_item,pay_method,price,cost,vat_yn,cost_vat,customer,customer_no,pay_yn,slip_memo,tax_bill_yn,cancel_yn,end_yn,reg_id,reg_user,reg_date) values ('"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&company&"','"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&pay_yn&"','"&slip_memo&"','Y','N','"&end_yn&"','"&user_id&"','"&user_name&"',now())"
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
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
