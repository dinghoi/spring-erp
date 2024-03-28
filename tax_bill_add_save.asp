<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	slip_seq = request.form("slip_seq")
	slip_date = request.form("slip_date")
	bonbu = request.form("bonbu")
	saupbu = request.form("saupbu")
	team = request.form("team")
	org_name = reside_place
	accountitem = request.form("account")
	i=instr(1,accountitem,"/")'
	account = mid(accountitem,1,i-1)
	account_item = mid(accountitem,i+1)
	pay_method = "현금"
	paper_no = request.form("paper_no")
	price = int(request.form("price"))
	cost = int(request.form("cost"))
	cost_vat = int(request.form("cost_vat"))
	vat_yn = "Y"
	customer = request.form("customer")
	pay_yn = "N"
	slip_memo = request.form("slip_memo")
	end_yn = request.form("end_yn")
	slip_gubun = "계산서"

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select * from customer where customer_no = '" + customer + "'"
	set rs=dbconn.execute(sql)
  	customer = rs("customer")
	customer_no = rs("customer_no")
	rs.close()

	if	u_type = "U" then
		sql = "update general_cost set team='"&team&"',org_name='"&org_name&"',account='"&account&"',account_item='"&account_item&"',pay_method='"&pay_method&"',price="&price&",cost="&cost&",vat_yn='"&vat_yn&"',cost_vat="&cost_vat&",customer='"&customer&"',paper_no='"&paper_no&"',use_man='"&use_man&"',emp_no='"&emp_no&"',pay_yn='"&pay_yn&"',slip_memo='"&slip_memo&"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where slip_date='"&slip_date&"' and slip_seq = '"&slip_seq&"'"
		dbconn.execute(sql)	  
	  else
		sql="select max(slip_seq) as max_seq from general_cost where slip_date='" + slip_date + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			slip_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			slip_seq = right(max_seq,3)
		end if

		sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,bonbu,saupbu,team,org_name,account,account_item,pay_method,price,cost,vat_yn,cost_vat,customer,customer_no"
		sql = sql +	",paper_no,pay_yn,slip_memo,end_yn,reg_id,reg_user,reg_date) values "
		sql = sql +	" ('"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&paper_no&"','"&pay_yn&"','"&slip_memo&"','"&end_yn&"','"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)
	end if

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
