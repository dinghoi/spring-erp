<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	edu_seq = request.form("edu_seq")
	edu_empno = request.form("edu_empno")
	
	'response.write edu_empno
	'response.write"alert('"&edu_empno&"');"
	
	edu_start_date = request.form("edu_start_date")
    edu_end_date = request.form("edu_end_date")
    edu_name = request.form("edu_name")
	edu_office = request.form("edu_office")
    edu_finish_no = request.form("edu_finish_no")
	edu_pay = 0
    'edu_pay = request.form("edu_pay")
    edu_comment = request.form("edu_comment")
    'edu_reg_date = request.form("edu_reg_date")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_edu set edu_name='"&edu_name&"',edu_office='"&edu_office&"',edu_finish_no='"&edu_finish_no&"',edu_start_date='"&edu_start_date&"',edu_end_date='"&edu_end_date&"',edu_comment='"&edu_comment&"',edu_mod_date=now(),edu_mod_user='"&emp_user&"' where edu_empno ='"&edu_empno&"' and edu_seq = '"&edu_seq&"'"
		dbconn.execute(sql)	  
	  else
		sql="select max(edu_seq) as max_seq from emp_edu where edu_empno='" + edu_empno + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			edu_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			edu_seq = right(max_seq,3)
		end if

		sql = "insert into emp_edu (edu_empno,edu_seq,edu_name,edu_office,edu_finish_no,edu_start_date,edu_end_date,edu_pay,edu_comment,edu_reg_date,edu_reg_user) values "
		sql = sql +	" ('"&edu_empno&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"','"&edu_start_date&"','"&edu_end_date&"','"&edu_pay&"','"&edu_comment&"',now(),'"&emp_user&"')"
		

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
