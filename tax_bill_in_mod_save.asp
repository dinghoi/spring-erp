<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'	on Error resume next

	dbconn.BeginTrans

	slip_date = request.form("slip_date")
	slip_seq = request.form("slip_seq")
	slip_gubun = request.form("slip_gubun")
	bonbu = request.form("bonbu")
	saupbu = request.form("saupbu")
	team = request.form("team")
	org_name = request.form("org_name")
	reside_place = request.form("reside_place")
	if isnull(reside_place) then
		reside_place = ""
	end if
	emp_no = request.form("emp_no")
	sql="select * from emp_master where emp_no='"&emp_no&"'"
	set rs_emp=dbconn.execute(sql)
	emp_grade = rs_emp("emp_job")
	emp_name = rs_emp("emp_name")
	company = request.form("company")
	account = request.form("account")
	account_item = request.form("account_item")
	slip_memo = request.form("slip_memo")
	mg_saupbu = request.form("mg_saupbu") 
	pl_yn = request.form("pl_yn") 
	
	sql = "Update general_cost set slip_gubun='"&slip_gubun&"',bonbu='"&bonbu&"',saupbu='"&saupbu&"',team='"&team&"',org_name='"&org_name&"',reside_place='"&reside_place&"',company='"&company&"',emp_name='"&emp_name&"',emp_no='"&emp_no&"',emp_grade='"&emp_grade&"',account='"&account&"',account_item='"&account_item&"',slip_memo='"&slip_memo&"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now(),mg_saupbu = '"&mg_saupbu&"',pl_yn = '"&pl_yn&"' where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
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
