<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	old_order_no = request.form("old_order_no")
	old_order_seq = request.form("old_order_seq")
	old_order_date = request.form("old_order_date")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "delete from met_order where (order_no = '"&old_order_no&"') and (order_seq = '"&old_order_seq&"') and (order_date = '"&old_order_date&"')"
	dbconn.execute(sql)
	sql = "delete from met_order_goods where (og_order_no = '"&old_order_no&"') and (og_order_seq = '"&old_order_seq&"') and (og_order_date = '"&old_order_date&"')"
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
