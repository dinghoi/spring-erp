<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	trade_code = request.form("trade_code")
	group_name = request.form("group_name")
	bill_trade_name = request.form("bill_trade_name")
	emp_no = request.form("emp_no")
	emp_name = request.form("emp_name")
	saupbu = request.form("saupbu")
	trade_id = request.form("trade_id")
	use_sw = request.form("use_sw")
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	Sql="select * from sales_org where saupbu = '"&saupbu&"'"
	Set rs=DbConn.Execute(Sql)
	if rs.eof or rs.bof then
		saupbu = ""
	end if

	sql = "Update trade set bill_trade_name='"&bill_trade_name&"',trade_id ='"&trade_id&"',emp_no='"&emp_no&"',emp_name='"&emp_name&"',saupbu='"&saupbu&"',group_name='"&group_name&"',use_sw='"&use_sw&"',mod_id='"&user_id&"',mod_date=now() where trade_code ='"&trade_code&"'"
	dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "처리 되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"opener.document.frm.submit();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
