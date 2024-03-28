<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
	on Error resume next

	response.write("시작")
	emp_no = 1000000
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select * from memb order by user_name asc"
	Rs.Open Sql, Dbconn, 1

	do until rs.eof
		emp_no = emp_no + 1
		sql = "update memb set emp_no='"&cstr(emp_no)&"'where user_id='"&rs("user_id")&"'"
		dbconn.execute(sql)	  
		rs.movenext()
	loop

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
