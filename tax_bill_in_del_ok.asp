<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	slip_seq = request.form("slip_seq")
	slip_date = request.form("slip_date")
	approve_no = request.form("approve_no")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "delete from general_cost where slip_date ='"&slip_date&"' and slip_seq='"&slip_seq&"'"
	dbconn.execute(sql)

	sql = "Update tax_bill set cost_reg_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where approve_no = '"&approve_no&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
