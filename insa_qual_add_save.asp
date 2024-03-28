<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	qual_seq = request.form("qual_seq")
	qual_empno = request.form("qual_empno")

	qual_type = request.form("qual_type")
    qual_grade = request.form("qual_grade")
    qual_pass_date = request.form("qual_pass_date")
    qual_org = request.form("qual_org")
    qual_no = request.form("qual_no")
	qual_passport = request.form("qual_passport")
	qual_pay_id = request.form("qual_pay_id")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_qual set qual_type='"&qual_type&"',qual_grade='"&qual_grade&"',qual_pass_date='"&qual_pass_date&"',qual_org='"&qual_org&"',qual_no='"&qual_no&"',qual_passport='"&qual_passport&"',qual_pay_id='"&qual_pay_id&"',qual_mod_date=now(),qual_mod_user='"&emp_user&"' where qual_empno ='"&qual_empno&"' and qual_seq = '"&qual_seq&"'"
		dbconn.execute(sql)	  
	  else
		sql="select max(qual_seq) as max_seq from emp_qual where qual_empno='" + qual_empno + "'"
		
		'response.write sql
		
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			qual_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			qual_seq = right(max_seq,3)
		end if

		sql = "insert into emp_qual(qual_empno,qual_seq,qual_type,qual_grade,qual_pass_date,qual_org,qual_no,qual_passport,qual_pay_id,qual_reg_date,qual_reg_user) values "
		sql = sql +	" ('"&qual_empno&"','"&qual_seq&"','"&qual_type&"','"&qual_grade&"','"&qual_pass_date&"','"&qual_org&"','"&qual_no&"','"&qual_passport&"','"&qual_pay_id&"',now(),'"&emp_user&"')"
		
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
