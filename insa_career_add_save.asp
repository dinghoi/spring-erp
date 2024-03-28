<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	career_seq = request.form("career_seq")
	career_empno = request.form("career_empno")
	
    career_join_date = request.form("career_join_date")
    career_end_date = request.form("career_end_date")
    career_office = request.form("career_office")
    career_dept = request.form("career_dept")
    career_position = request.form("career_position")
    career_task = request.form("career_task")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_career set career_join_date='"&career_join_date&"',career_end_date='"&career_end_date&"',career_office='"&career_office&"',career_dept='"&career_dept&"',career_position='"&career_position&"',career_task='"&career_task&"',career_mod_date= now(),career_mod_user='"&emp_user&"' where career_empno ='"&career_empno&"' and career_seq = '"&career_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(career_seq) as max_seq from emp_career where career_empno='" + career_empno + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			career_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			career_seq = right(max_seq,3)
		end if

		sql = "insert into emp_career(career_empno,career_seq,career_join_date,career_end_date,career_office,career_dept,career_position,career_task,career_reg_date,career_reg_user) values "
		sql = sql +	" ('"&career_empno&"','"&career_seq&"','"&career_join_date&"','"&career_end_date&"','"&career_office&"','"&career_dept&"','"&career_position&"','"&career_task&"',now(),'"&emp_user&"')"
		
		'response.write sql
		
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
