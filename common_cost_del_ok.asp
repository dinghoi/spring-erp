<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'on Error resume next

u_type = request.form("u_type")
slip_seq = request.form("slip_seq")
slip_date = request.form("slip_date")
old_date = request.form("old_date")

set dbconn = server.CreateObject("adodb.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

dbconn.BeginTrans

sql = "delete from general_cost where slip_date ='"&old_date&"' and slip_seq='"&slip_seq&"'"
dbconn.execute(sql)

if Err.number <> 0 then
	dbconn.RollbackTrans
	end_msg = "삭제 중 Error가 발생하였습니다."
else
	dbconn.CommitTrans
	end_msg = "삭제되었습니다."
end if

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>
