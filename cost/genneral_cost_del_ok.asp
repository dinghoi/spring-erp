<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

u_type = request.form("u_type")
slip_seq = request.form("slip_seq")
slip_date = request.form("slip_date")
old_date = request.form("old_date")

'Response.write old_date & " " & slip_seq
'Response.end

set dbconn = server.CreateObject("adodb.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

dbconn.BeginTrans

SQL = "delete from general_cost where slip_date ='"&old_date&"' and slip_seq='"&slip_seq&"' "
dbconn.execute(sql)

if Err.number <> 0 then
	dbconn.RollbackTrans
	end_msg = "삭제중 Error가 발생하였습니다."
else
	dbconn.CommitTrans
	end_msg = "삭제되었습니다."
end if

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.write "	self.opener.location.reload();"
Response.write "	window.close();"
Response.write"</script>"
Response.End

dbconn.Close() : Set dbconn = Nothing

%>
