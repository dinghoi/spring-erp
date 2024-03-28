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
	mg_group = request.form("mg_group")
	etc_amt = int(request.form("etc_amt"))
	used_sw = request.form("used_sw")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	Sql="select * from type_code where etc_type = '"&etc_type&"'"
	response.write(sql)
	Set rs=DbConn.Execute(Sql)
	if rs.eof or rs.bof then
		type_name = "error"
	  else
	  	type_name = rs("type_name")
	end if
	
	if	u_type = "U" then
		sql = "Update etc_code set etc_name='"&etc_name&"',etc_group='"&etc_group&"',group_name ='"&group_name& _
		"',mg_group ='0',etc_amt ="&etc_amt&",used_sw='"&used_sw&"',reg_id='"&user_id&"',reg_date=now() where etc_code = '"&etc_code&"'"
		dbconn.execute(sql)
	  else
		sql="insert into etc_code (etc_code,etc_type,type_name,etc_name,etc_group,group_name,mg_group,etc_amt,used_sw,reg_id,reg_date) "& _
		"values ('"&etc_code&"','"&etc_type&"','"&type_name&"','"&etc_name&"','"&etc_group&"','"&group_name&"','0',"&etc_amt& _
		",'"&used_sw&"','"&user_id&"',now())"
		dbconn.execute(sql)
	end if	

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.Redirect "account_etc_code_mg.asp?etc_type="&etc_type
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
