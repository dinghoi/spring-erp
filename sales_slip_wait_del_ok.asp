<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	old_slip_id = abc("old_slip_id")
	old_slip_no = abc("old_slip_no")
	old_slip_seq = abc("old_slip_seq")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "delete from sales_slip where slip_id ='"&old_slip_id&"' and slip_no='"&old_slip_no&"' and slip_seq='"&old_slip_seq&"'"
	dbconn.execute(sql)
	sql = "delete from sales_slip_detail where slip_id ='"&old_slip_id&"' and slip_no='"&old_slip_no&"' and slip_seq='"&old_slip_seq&"'"
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
