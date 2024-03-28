<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	s_id = request.form("s_id")
	
	s_emp_no = request.form("s_emp_no")
	s_year = request.form("s_year")
	s_seq = request.form("s_seq")
	
	s_emp_name = request.form("s_emp_name")
	s_type = request.form("s_type")
	s_bank_code = request.form("s_bank_code")
	s_bank_name = request.form("s_bank_name")
	s_account_no = request.form("s_account_no")

	s_amt =int(request.form("s_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_saving set s_type='"&s_type&"',s_bank_code='"&s_bank_code&"',s_bank_name='"&s_bank_name&"',s_account_no='"&s_account_no&"',s_amt='"&s_amt&"' where s_year ='"&s_year&"' and s_emp_no = '"&s_emp_no&"' and s_id = '"&s_id&"' and s_seq = '"&s_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(s_seq) as max_seq from pay_yeartax_saving where s_year='" + s_year + "' and s_emp_no='" + s_emp_no + "' and s_id='" + s_id + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			s_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			s_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_saving (s_year,s_emp_no,s_id,s_seq,s_emp_name,s_type,s_bank_code,s_bank_name,s_account_no,s_amt) values "
		sql = sql +	" ('"&s_year&"','"&s_emp_no&"','"&s_id&"','"&s_seq&"','"&s_emp_name&"','"&s_type&"','"&s_bank_code&"','"&s_bank_name&"','"&s_account_no&"','"&s_amt&"')"
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
