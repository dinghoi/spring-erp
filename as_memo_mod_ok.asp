<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	acpt_no = request.form("acpt_no")
	as_memo = request.form("as_memo")
	as_memo = Replace(as_memo,"'","&quot;")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "Update as_acpt set as_memo='"&as_memo&"', mod_id='"&user_id&"', mod_date=now() where acpt_no = "&int(acpt_no)
	dbconn.execute(sql)

	mod_pg = "장애내용"
	sql = "insert into nkp.as_mod (acpt_no,mod_date,mod_id,mod_name,mod_pg) values ('"&acpt_no&"',now(),'"&user_id&"','"&user_name&"','"&mod_pg&"')"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
