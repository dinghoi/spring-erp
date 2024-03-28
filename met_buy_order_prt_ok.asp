<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	old_order_no = request.form("old_order_no")
	old_order_seq = request.form("old_order_seq")
	old_order_date = request.form("old_order_date")
	
	order_buy_no = request.form("order_buy_no")
	order_buy_seq = request.form("order_buy_seq")
	order_buy_date = request.form("order_buy_date")
	
	order_ing = "3"
	buy_ing = "3"

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

'구매요청 update
    sql = "Update met_buy set buy_ing='"&buy_ing&"',mod_date=now(),mod_user='"&user_name&"' where buy_no = '"&order_buy_no&"' and buy_no = '"&order_buy_seq&"' and buy_date = '"&order_buy_date&"'"
	dbconn.execute(sql)
	
'발주 update
    sql = "Update met_order set order_ing='"&order_ing&"',mod_date=now(),mod_user='"&user_name&"' where (order_no = '"&old_order_no&"') and (order_seq = '"&old_order_seq&"') and (order_date = '"&old_order_date&"')"
	dbconn.execute(sql)	

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "출력 되었습니다...."
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
