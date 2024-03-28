-9<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	u_type = request.form("u_type")
	etc_code = request.form("etc_code")
	etc_type = "51"
	type_name = request.form("type_name")
	etc_name = request.form("etc_name")
	group_name = request.form("group_name")
	used_sw = request.form("used_sw")
	etc_amt = 0

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	if	u_type = "U" then
		sql = "Update etc_code set etc_name='"&etc_name&"',group_name ='"&group_name& _
		"',mg_group ='0',etc_amt =0,used_sw='"&used_sw&"',reg_id='"&user_id&"',reg_date=now() where etc_code = '"&etc_code&"'"
		dbconn.execute(sql)
	  else
		sql="select max(etc_code) as max_no from etc_code where etc_type = '51'"
		set rs=dbconn.execute(sql)

		if	isnull(rs("max_no"))  then
			etc_code = "5101"
		  else
			etc_code = cstr(int(rs("max_no")) + 1)
		end if
		slip_seq = "00"
		sql="insert into etc_code (etc_code,etc_type,type_name,etc_name,group_name,mg_group,etc_amt,used_sw,reg_id,reg_date) "& _
		"values ('"&etc_code&"','"&etc_type&"','"&type_name&"','"&etc_name&"','"&group_name&"','0',"&etc_amt& _
		",'"&used_sw&"','"&user_id&"',now())"
		dbconn.execute(sql)
	end if

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"
	response.Redirect "sales_goods_code_mg.asp?type_name="&type_name
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing


%>
