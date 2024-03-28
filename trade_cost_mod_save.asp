<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	trade_code      = request.form("trade_code")
	group_name      = request.form("group_name")
	bill_trade_name = request.form("bill_trade_name")
	emp_no          = request.form("emp_no")
	emp_name        = request.form("emp_name")
	saupbu          = request.form("saupbu")
	trade_id        = request.form("trade_id")
	use_sw          = request.form("use_sw")

	cost_year = year(now())

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans


'Response.write "<pre>"&saupbu&"</pre>"

	Sql = "SELECT *                               " & chr(13) &_
	      "  FROM sales_org                       " & chr(13) &_
	      " WHERE sales_year = '" &cost_year & "' " & chr(13) &_
	      "   AND saupbu     = '"&saupbu&"'       " & chr(13) &_
	      " LIMIT 1                               "

'Response.write "<pre>"&Sql&"</pre>"

	Set rs=DbConn.Execute(Sql)
	if rs.eof or rs.bof then
		saupbu = ""
	end if

	sql = " UPDATE trade                                    " & chr(13) &_
		  "    SET bill_trade_name = '"&bill_trade_name&"'  " & chr(13) &_
		  "       ,trade_id        = '"&trade_id&"'         " & chr(13) &_
		  "       ,emp_no          = '"&emp_no&"'           " & chr(13) &_
		  "       ,emp_name        = '"&emp_name&"'         " & chr(13) &_
		  "       ,saupbu          = '"&saupbu&"'           " & chr(13) &_
		  "       ,group_name      = '"&group_name&"'       " & chr(13) &_
		  "       ,use_sw          = '"&use_sw&"'           " & chr(13) &_
		  "       ,mod_id          = '"&user_id&"'          " & chr(13) &_
		  "       ,mod_date        = now()                  " & chr(13) &_
		  "  WHERE trade_code = '"&trade_code&"'            "

'Response.write "<pre>"&Sql&"</pre>"
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
