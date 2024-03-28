<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	approve_no = request.form("approve_no")
	cancel_yn = request.form("cancel_yn")
	cost = int(request.form("cost"))
	cost_vat = int(request.form("cost_vat"))
	account = request.form("account")
	account_item = request.form("account_item")
	emp_no = request.form("emp_no")
	old_emp_no = request.form("old_emp_no")
	pl_yn = request.form("pl_yn")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if old_emp_no <> emp_no then
		Sql="select * from memb where user_id = '"&emp_no&"'"
		Set rs_memb=DbConn.Execute(Sql)
		emp_name = rs_memb("user_name")
		emp_company = rs_memb("emp_company")
		bonbu = rs_memb("bonbu")
		saupbu = rs_memb("saupbu")
		team = rs_memb("team")
		org_name = rs_memb("org_name")
		reside_place = rs_memb("reside_place")
		reside_company = rs_memb("reside_company")

		sql = "Update card_slip set emp_no='"&emp_no&"',emp_name='"&emp_name&"',emp_company='"&emp_company&"',saupbu='"&saupbu&"',bonbu='"&bonbu&"',team='"&team&"',org_name='"&org_name&"',reside_place='"&reside_place&"',reside_company='"&reside_company&"',cost='"&cost&"',cost_vat='"&cost_vat&"',account='"&account&"',account_item='"&account_item&"',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now(),pl_yn='"&pl_yn&"' where approve_no = '"&approve_no&"' and cancel_yn ='"&cancel_yn&"'"
		dbconn.execute(sql)
	  else	  
		sql = "Update card_slip set cost='"&cost&"',cost_vat='"&cost_vat&"',account='"&account&"',account_item='"&account_item&"',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now(),pl_yn='"&pl_yn&"' where approve_no = '"&approve_no&"' and cancel_yn ='"&cancel_yn&"'"
		dbconn.execute(sql)
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "변경되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
