<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	bill_issue_date = request.form("bill_issue_date")
	slip_id = request.form("slip_id")
	slip_no = request.form("slip_no")
	slip_seq = request.form("slip_seq")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans
	
	sql = "update sales_slip set bill_issue_date='"&bill_issue_date&"', mod_emp_no='"&user_id&"', mod_name='"&user_name&"', mod_date=now() where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
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
	response.write"opener.document.frm.submit();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
