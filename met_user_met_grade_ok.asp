<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	view_condi = request.form("view_condi1")

    met_grade = request.form("met_grade")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_mem = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans
	
	sql = "Update memb set met_grade='"&met_grade&"' where user_id='"&view_condi&"'"
	dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "변경되었습니다...."
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
	