<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'	on Error resume next

	dbconn.BeginTrans

	approve_no = request.form("approve_no")
	Sql="select * from tax_bill where approve_no = '"&approve_no&"'"
	Set rs=DbConn.Execute(Sql)

	Sql="select * from trade where trade_no = '"&rs("trade_no")&"'"
	Set rs_trade=DbConn.Execute(Sql)
	if rs_trade.eof or rs_trade.bof then
		customer = rs("trade_name")
	  else
		customer = rs_trade("trade_name")
	end if

	slip_date = rs("bill_date")
	slip_gubun = request.form("slip_gubun")
	company = request.form("company")
	emp_company = rs("owner_company")
	bonbu = request.form("bonbu")
	saupbu = request.form("saupbu")
	if isnull(saupbu) then
		saupbu = ""
	end if
	team = request.form("team")
	org_name = request.form("org_name")
	reside_place = request.form("reside_place")
	if isnull(reside_place) then
		reside_place = ""
	end if
	account = request.form("account")
	account_item = request.form("account_item")
	pay_method = "현금"

	'price = int(rs("price"))
	'cost = int(rs("cost"))
	'cost_vat = int(rs("cost_vat"))

	price = cdbl(rs("price"))
	cost = cdbl(rs("cost"))
	cost_vat = cdbl(rs("cost_vat"))

	vat_yn = "Y"
	customer = rs("trade_name")
	customer_no = rs("trade_no")
	pay_yn = "N"
	slip_memo = request.form("slip_memo")
	end_yn = request.form("end_yn")
	mg_saupbu = request.form("mg_saupbu")

	emp_no = request.form("emp_no")
	sql="select * from emp_master where emp_no='"&emp_no&"'"
	set rs_emp=dbconn.execute(sql)
	emp_grade = rs_emp("emp_job")
	emp_name = rs_emp("emp_name")

	sql="select max(slip_seq) as max_seq from general_cost where slip_date='"&slip_date&"'"
	set rs=dbconn.execute(sql)

	if	isnull(rs("max_seq"))  then
		slip_seq = "001"
	  else
		max_seq = "00" + cstr((int(rs("max_seq")) + 1))
		slip_seq = right(max_seq,3)
	end if

	' 2019.02.02 [박성민 요청] "하장호"의 일반경비 등록시 트랜젹션문제로 ISERT는 되나 UPDATE가 안돼는 문제발생..	트랜잭션 TEST해볼것...
	sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,emp_company,bonbu,saupbu,team,org_name,reside_place,company,account,account_item"&",pay_method,price,cost,vat_yn,cost_vat,customer,customer_no,emp_name,emp_no,emp_grade,pay_yn,slip_memo,tax_bill_yn,cancel_yn,end_yn,reg_id,reg_user,reg_date,approve_no,mg_saupbu) values "&"('"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&company&"','"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&emp_name&"','"&emp_no&"','"&emp_grade&"','"&pay_yn&"','"&slip_memo&"','Y','N','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&approve_no&"','"&mg_saupbu&"')"

	dbconn.execute(sql)

	sql = "Update tax_bill set cost_reg_yn='Y',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where approve_no = '"&approve_no&"'"
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
