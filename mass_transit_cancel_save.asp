<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	cancel_yn = request.form("cancel_yn")	
	mg_ce_id = request.form("mg_ce_id")	
	run_date = request.form("run_date")	
	run_seq = request.form("run_seq")	
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans


	sql = "update transit_cost set cancel_yn='"&cancel_yn&"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where mg_ce_id='"&mg_ce_id&"' and run_date = '"&run_date&"' and run_seq = '"&run_seq&"'"
	dbconn.execute(sql)
		
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
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
