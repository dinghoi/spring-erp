<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	cost_id=Request("cost_id")
	end_month=Request("end_month")
	end_yn=Request("end_yn")
	
	from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	
	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_into = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Set RsCount = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

	if cost_id = "야특근" then
		sql = "select * from overtime where work_date >= '" + from_date  + "' and work_date <= '" + to_date  + "'"
	end if	
	Rs.Open Sql, Dbconn, 1

	do until rs.eof
		if cost_id = "야특근" then
			sql = "Update overtime set end_yn='Y' where work_date = '"&rs("work_date")&"' and mg_ce_id = '"&rs("mg_ce_id")&"'"
			dbconn.execute(sql)
		end if
		rs.movenext()
	loop

	if end_yn = "C" then
		sql = "Update cost_end set end_yn='Y',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where cost_id = '"&cost_id& _
		"' and end_month = '"&end_month&"'"
	  else
		sql="insert into cost_end (cost_id,end_month,end_yn,reg_id,reg_name,reg_date) values ('"&cost_id&"','"&end_month& _
		"','Y','"&user_id&"','"&user_name&"',now())"
	end if
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
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


