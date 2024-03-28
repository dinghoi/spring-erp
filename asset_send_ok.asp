<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	
	asset_no = request.form("asset_no")
	serial_no = request.form("serial_no")
	dept_code = request.form("dept_code")
	user_name = request.form("user_name")
	send_date = request.form("send_date")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	
	sql = "Update asset set serial_no ='"+serial_no+"', dept_code ='"+dept_code+"', user_name ='"+user_name+"', inst_process = 'S' , send_date='"+send_date+"', mod_id='"+user_id+"', mod_date=now() where asset_no = '" + asset_no + "'"
	dbconn.execute(sql)

	end_msg = "발송 완료되었습니다...."
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>

