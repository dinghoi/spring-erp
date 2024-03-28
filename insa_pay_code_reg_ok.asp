<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	u_type = request.form("u_type")
	etc_code = request.form("etc_code")
	etc_type = request.form("etc_type")
	type_name = request.form("type_name")
	etc_name = request.form("etc_name")
	etc_group = request.form("etc_group")
	group_name = request.form("group_name")
	used_sw = request.form("used_sw")
	emp_tax_id = request.form("emp_tax_id")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	if	u_type = "U" then
		sql = "Update emp_etc_code set emp_etc_name='"&etc_name&"',emp_etc_group='"&etc_group&"',emp_group_name ='"&group_name&"',emp_used_sw='"&used_sw&"',emp_tax_id='"&emp_tax_id&"' where emp_etc_code = '"&etc_code&"'"
		dbconn.execute(sql)
	  else
		sql="insert into emp_etc_code (emp_etc_code,emp_etc_type,emp_type_name,emp_etc_name,emp_etc_group,emp_group_name,emp_mg_group,emp_used_sw,emp_tax_id) "
		sql=sql + "values ('"&etc_code&"','"&etc_type&"','"&type_name&"','"&etc_name&"','"&etc_group&"','"&group_name&"','"&mg_group&"','"&used_sw&"','"&emp_tax_id&"')"
		dbconn.execute(sql)
	end if

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"
	response.Redirect "/insa_pay_code_mg.asp?emp_etc_type="&etc_type
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing


%>
