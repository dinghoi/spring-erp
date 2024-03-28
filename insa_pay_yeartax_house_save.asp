<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	h_year = request.form("inc_yyyy")
	h_emp_no = request.form("emp_no")
	h_emp_name = request.form("emp_name")
	h_person_no = request.form("emp_person")

	'response.write(y_emp_no)
	'response.End

	h_lender_amt =int(request.form("h_lender_amt"))
	h_person_amt =int(request.form("h_person_amt"))
	h_month_amt =int(request.form("h_month_amt"))
	h_long15_amt =int(request.form("h_long15_amt"))
	h_long29_amt =int(request.form("h_long29_amt"))
	h_long30_amt =int(request.form("h_long30_amt"))
	h_fixed_amt =int(request.form("h_fixed_amt"))
	h_other_amt =int(request.form("h_other_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_house set h_lender_amt='"&h_lender_amt&"',h_person_amt='"&h_person_amt&"',h_long15_amt='"&h_long15_amt&"',h_long29_amt='"&h_long29_amt&"',h_long30_amt='"&h_long30_amt&"',h_fixed_amt='"&h_fixed_amt&"',h_other_amt='"&h_other_amt&"' where h_year ='"&h_year&"' and h_emp_no = '"&h_emp_no&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into pay_yeartax_house (h_year,h_emp_no,h_emp_name,h_person_no,h_lender_amt,h_person_amt,h_long15_amt,h_long29_amt,h_long30_amt,h_fixed_amt,h_other_amt) values "
		sql = sql +	" ('"&h_year&"','"&h_emp_no&"','"&h_emp_name&"','"&h_person_no&"','"&h_lender_amt&"','"&h_person_amt&"','"&h_long15_amt&"','"&h_long29_amt&"','"&h_long30_amt&"','"&h_fixed_amt&"','"&h_other_amt&"')"
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
	'response.write"self.opener.location.reload();"	
	response.write"location.replace('insa_pay_yeartax_house.asp');"	
	'response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
