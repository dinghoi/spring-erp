<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	b_emp_no = request.form("b_emp_no")
	b_year = request.form("b_year")
	b_seq = request.form("b_seq")
	
	b_emp_name = request.form("b_emp_name")
	b_company_no = request.form("b_company_no")
	b_company = request.form("b_company")
	b_from_date = request.form("b_from_date")
	b_to_date = request.form("b_to_date")

	b_pay =int(request.form("b_pay"))
	b_bonus =int(request.form("b_bonus"))
	b_deem_bonus =int(request.form("b_deem_bonus"))
	b_overtime_taxno =int(request.form("b_overtime_taxno"))
	b_nps =int(request.form("b_nps"))
	b_nhis =int(request.form("b_nhis"))
	b_epi =int(request.form("b_epi"))
	b_longcare =int(request.form("b_longcare"))
	b_income_tax =int(request.form("b_income_tax"))
	b_wetax =int(request.form("b_wetax"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_before set b_company_no='"&b_company_no&"',b_company='"&b_company&"',b_from_date='"&b_from_date&"',b_to_date='"&b_to_date&"',b_pay='"&b_pay&"',b_bonus='"&b_bonus&"',b_deem_bonus='"&b_deem_bonus&"',b_overtime_taxno='"&b_overtime_taxno&"',b_nps='"&b_nps&"',b_nhis='"&b_nhis&"',b_epi='"&b_epi&"',b_longcare='"&b_longcare&"',b_income_tax='"&b_income_tax&"',b_wetax='"&b_wetax&"' where b_year ='"&b_year&"' and b_emp_no = '"&b_emp_no&"' and b_seq = '"&b_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(b_seq) as max_seq from pay_yeartax_before where b_year='" + b_year + "' and b_emp_no='" + b_emp_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			b_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			b_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_before (b_year,b_emp_no,b_seq,b_emp_name,b_company_no,b_company,b_from_date,b_to_date,b_pay,b_bonus,b_deem_bonus,b_overtime_taxno,b_nps,b_nhis,b_epi,b_longcare,b_income_tax,b_wetax) values "
		sql = sql +	" ('"&b_year&"','"&b_emp_no&"','"&b_seq&"','"&b_emp_name&"','"&b_company_no&"','"&b_company&"','"&b_from_date&"','"&b_to_date&"','"&b_pay&"','"&b_bonus&"','"&b_deem_bonus&"','"&b_overtime_taxno&"','"&b_nps&"','"&b_nhis&"','"&b_epi&"','"&b_longcare&"','"&b_income_tax&"','"&b_wetax&"')"
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
