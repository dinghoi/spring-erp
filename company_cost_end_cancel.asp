<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	end_month=Request("end_month")

	cost_year = mid(end_month,1,4)
	cost_month = mid(end_month,5)
	
	from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
	'response.write(end_date)
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	
	response.write"<script language=javascript>"
	response.write"alert('마감 취소중!!!');"
	response.write"</script>"

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans
	
	sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '상주비용'"
	dbconn.execute(sql)

	sql = "update company_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
	Response.write sql
	dbconn.execute(sql)
	
	sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
	dbconn.execute(sql)
	sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "마감이 취소되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


