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

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_type = Server.CreateObject("ADODB.Recordset")	
	Dbconn.open dbconnect
	
	Sql="select * from met_type_code where etc_type = '" + etc_type + "'"
	Rs_type.Open Sql, Dbconn, 1
	if not Rs_type.eof then
	       type_name = Rs_type("type_name")
	   else
	       type_name = ""
	end if

	if	u_type = "U" then
		sql = "Update met_etc_code set etc_name='"&etc_name&"',etc_group='"&etc_group&"',group_name ='"&group_name&"',used_sw='"&used_sw&"' where etc_type = '" + etc_type + "' and etc_code = '"&etc_code&"'"
		dbconn.execute(sql)
	  else
		sql="insert into met_etc_code (etc_code,etc_type,type_name,etc_name,etc_group,group_name,mg_group,used_sw) "
		sql=sql + "values ('"&etc_code&"','"&etc_type&"','"&type_name&"','"&etc_name&"','"&etc_group&"','"&group_name&"','"&mg_group&"','"&used_sw&"')"
		dbconn.execute(sql)
	end if	

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.Redirect "met_control_code_mg.asp?etc_type="&etc_type
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
