<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	approve_no = request.form("approve_no")
	saupbu = request.form("saupbu")
	emp_name = request.form("emp_name")
	emp_no = request.form("emp_no")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "update saupbu_sales set saupbu='"&saupbu&"', emp_name='"&emp_name&"', emp_no='"&emp_no&"', reg_id='"&user_id&"', reg_name='"&user_name&"', reg_date=now() where approve_no='"&approve_no&"'"
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
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
