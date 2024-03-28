<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	
	asset_no = request.form("asset_no")
	user_name = request.form("user_name")
	inst_process = request.form("inst_process")
	install_date = request.form("install_date")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	
	if inst_process = "Y" then
		sql = "Update asset set user_name ='"+user_name+"', inst_process='"+inst_process+"', install_date='"+install_date+"', mod_id='"+user_id+"', mod_date=now() where asset_no = '" + asset_no + "'"
	  else
		sql = "Update asset set user_name ='"+user_name+"', inst_process='"+inst_process+"', return_date='"+install_date+"', mod_id='"+user_id+"', mod_date=now() where asset_no = '" + asset_no + "'"
	end if
	dbconn.execute(sql)

	end_msg = "처리 완료되었습니다...."
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>

