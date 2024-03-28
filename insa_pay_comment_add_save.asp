<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	pmg_emp_no = request.form("pmg_emp_no")
	pmg_company = request.form("pmg_company")
	pmg_yymm = request.form("pmg_yymm")
	pmg_comment = request.form("pmg_comment")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_give = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

    emp_user = request.cookies("nkpmg_user")("coo_user_name")

		sql = "Update pay_month_give set pmg_comment='"&pmg_comment&"',pmg_mod_user='"&emp_user&"',pmg_mod_date=now() where pmg_yymm = '"&pmg_yymm&"' and pmg_id = '1' and pmg_emp_no = '"&pmg_emp_no&"' and pmg_company = '"&pmg_company&"'"
		
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
