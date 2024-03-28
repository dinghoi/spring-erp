<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	empno = request.form("empno")
	commute_date = request.form("commute_date")

	commute_time = request.form("commute_time")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "S" then
		sql = "insert into commute(emp_no,wrkt_dt,wrk_start_time) values "
		sql = sql +	" ('"&empno&"','"&commute_date&"','"&commute_time&"') on duplicate key update wrk_start_time ='"&commute_time&"'"
		dbconn.execute(sql)	  
	else

		sql = "insert into commute(emp_no,wrkt_dt,wrk_end_time) values "
		sql = sql +	" ('"&empno&"','"&commute_date&"','"&commute_time&"') on duplicate key update wrk_end_time ='"&commute_time&"'"
		
		response.write sql
		
		dbconn.execute(sql)
	end if

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
