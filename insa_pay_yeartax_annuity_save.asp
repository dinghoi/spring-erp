<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	a_emp_no = request.form("a_emp_no")
	a_year = request.form("a_year")
	a_seq = request.form("a_seq")
	
	a_emp_name = request.form("a_emp_name")
	a_type = request.form("a_type")
	a_bank_code = request.form("a_bank_code")
	a_bank_name = request.form("a_bank_name")
	a_account_no = request.form("a_account_no")

	a_amt =int(request.form("a_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_annuity set a_type='"&a_type&"',a_bank_code='"&a_bank_code&"',a_bank_name='"&a_bank_name&"',a_account_no='"&a_account_no&"',a_amt='"&a_amt&"' where a_year ='"&a_year&"' and a_emp_no = '"&a_emp_no&"' and a_seq = '"&a_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(a_seq) as max_seq from pay_yeartax_annuity where a_year='" + a_year + "' and a_emp_no='" + a_emp_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			a_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			a_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_annuity (a_year,a_emp_no,a_seq,a_emp_name,a_type,a_bank_code,a_bank_name,a_account_no,a_amt) values "
		sql = sql +	" ('"&a_year&"','"&a_emp_no&"','"&a_seq&"','"&a_emp_name&"','"&a_type&"','"&a_bank_code&"','"&a_bank_name&"','"&a_account_no&"','"&a_amt&"')"
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
