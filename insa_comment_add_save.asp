<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	cmt_date = request.form("cmt_date")
	cmt_empno = request.form("cmt_empno")
	cmt_emp_name = request.form("emp_name")
	cmt_company = request.form("emp_company")
	cmt_bonbu = request.form("emp_bonbu")
	cmt_saupbu = request.form("emp_saupbu")
	cmt_team = request.form("emp_team")
	cmt_org_code = request.form("emp_org_code")
	cmt_org_name = request.form("emp_org_name")
	cmt_comment = request.form("cmt_comment")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_comment set cmt_comment='"&cmt_comment&"',cmt_reg_date= now(),cmt_reg_user='"&emp_user&"' where cmt_empno ='"&cmt_empno&"' and cmt_date = '"&cmt_date&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into emp_comment (cmt_empno,cmt_date,cmt_emp_name,cmt_company,cmt_bonbu,cmt_saupbu,cmt_team,cmt_org_name,cmt_org_code,cmt_comment,cmt_reg_date,cmt_reg_user) values "
		sql = sql +	" ('"&cmt_empno&"','"&cmt_date&"','"&cmt_emp_name&"','"&cmt_company&"','"&cmt_bonbu&"','"&cmt_saupbu&"','"&cmt_team&"','"&cmt_org_name&"','"&cmt_org_code&"','"&cmt_comment&"',now(),'"&emp_user&"')"
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
