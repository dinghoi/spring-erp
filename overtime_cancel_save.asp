<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	mg_ce_id = request.form("mg_ce_id")
	work_date = request.form("work_date")
	cancel_yn = request.form("cancel_yn")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "UPDATE overtime SET cancel_yn='"&cancel_yn&"',mod_id='"&user_id&"',mod_user='"&user_name& _
		"',mod_date=now() where work_date = '"&work_date&"' and mg_ce_id = '"&mg_ce_id&"'"
		dbconn.execute(sql)
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "처리중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
		end_msg = "처리 되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
