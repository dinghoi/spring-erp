<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	ce_id = request("ce_id")
	mod_ce_id = request("mod_ce_id")
	
	mod_id = user_id

	set dbconn = server.CreateObject("adodb.connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

	sql="select * from memb where user_id = '" + mod_ce_id + "'"

	Set rs=DbConn.Execute(sql)

	user_name = rs("user_name")
	c_grade = rs("grade")
	rs.close()

	
	sql = "update ce_area set mg_ce_id='"&mod_ce_id&"', mod_date=now(), mod_id='"&mod_id&"' where mg_group = '" + mg_group + "' and mg_ce_id='" + ce_id + "'"
	dbconn.execute(sql)

	sql = "update area_mg set mg_ce_id='"&mod_ce_id&"', mod_date=now(), mod_id='"&mod_id&"' where mg_group = '" + mg_group + "' and mg_ce_id='" + ce_id + "'"
	dbconn.execute(sql)

	sql = "update juso_list set mg_ce_id='"&mod_ce_id&"', regi_date=now(), regi_id='"&mod_id&"' where mg_group = '" + mg_group + "' and reside = '0' and mg_ce_id='" + ce_id + "'"
	dbconn.execute(sql)

	sql = "update as_acpt set mg_ce_id='"&mod_ce_id&"', mg_ce='"&user_name&"', mod_date=now(), mod_id='"&mod_id&"' where mg_group = '" + mg_group + "' and (as_process = '접수' or as_process = '입고' or as_process = '연기' or as_process = '대체입고') and mg_ce_id='" + ce_id + "'"
	dbconn.execute(sql)
				
	response.write"<script language=javascript>"
	response.write"alert('변경 완료 되었습니다....');"		
'	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
'	response.write"location.replace('k1_ce_mg_list.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>

