<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	emp_no = request.form("emp_no")
	emp_name = request.form("emp_name")
	person_no1 = request.form("person_no1")
	person_no2 = request.form("person_no2")
	bank_code = ""
	bank_name = request.form("bank_name")
	account_no = request.form("account_no")
	account_holder = request.form("account_holder")
	emp_type = ""
	emp_pay_type = ""

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_emp = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	
	Sql="select * from emp_etc_code where emp_etc_type = '50' and emp_etc_name = '"&bank_name&"'"
	Rs_etc.Open Sql, Dbconn, 1
	bank_code = rs_etc("emp_etc_code")
	rs_etc.close()

    Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
    Set rs_emp = DbConn.Execute(SQL)
	if not rs_emp.eof then
         emp_type = rs_emp("emp_type")
	     emp_pay_type = rs_emp("emp_pay_type")
	end if
    rs_emp.close()

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_bank_account set bank_code='"&bank_code&"',bank_name='"&bank_name&"',account_no='"&account_no&"',account_holder='"&account_holder&"',mod_date= now(),mod_user='"&emp_user&"' where emp_no ='"&emp_no&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into pay_bank_account (emp_no,emp_name,person_no1,person_no2,emp_type,emp_pay_type,bank_code,bank_name,account_no,account_holder,reg_date,reg_user) values "
		sql = sql +	" ('"&emp_no&"','"&emp_name&"','"&person_no1&"','"&person_no2&"','"&emp_type&"','"&emp_pay_type&"','"&bank_code&"','"&bank_name&"','"&account_no&"','"&account_holder&"',now(),'"&emp_user&"')"
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
