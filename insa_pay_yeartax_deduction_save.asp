<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	old_de_id = request.form("old_de_id")
	
	de_year = request.form("inc_yyyy")
	de_id = request.form("de_id")
	de_emp_no = request.form("emp_no")
	de_emp_name = request.form("emp_name")
	de_person_no = request.form("emp_person")
	young_fdate = request.form("young_fdate")
	young_ldate = request.form("young_ldate")
	de_tax_nation = request.form("de_tax_nation")
	de_tax_date = request.form("de_tax_date")
	de_report_date = request.form("de_report_date")
	de_office = request.form("de_office")
	de_stay = request.form("de_stay")
	de_position = request.form("de_position")

	'response.write(y_emp_no)
	'response.End

	de_wonchen =int(request.form("de_wonchen"))
	de_tax_s =int(request.form("de_tax_s"))
	de_tax_w =int(request.form("de_tax_w"))
	
	if young_fdate = "" or isnull(young_fdate) then
	       young_fdate = "1900-01-01"
	end if
	if young_ldate = "" or isnull(young_ldate) then
	       young_ldate = "1900-01-01"
	end if
	if de_tax_date = "" or isnull(de_tax_date) then
	       de_tax_date = "1900-01-01"
	end if
	if de_report_date = "" or isnull(de_report_date) then
	       de_report_date = "1900-01-01"
	end if
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_deduction set young_fdate='"&young_fdate&"',young_ldate='"&young_ldate&"',de_tax_s='"&de_tax_s&"',de_tax_w='"&de_tax_w&"',de_tax_nation='"&de_tax_nation&"',de_tax_date='"&de_tax_date&"',de_report_date='"&de_report_date&"',de_office='"&de_office&"',de_stay='"&de_stay&"',de_position='"&de_position&"' where de_year ='"&de_year&"' and de_emp_no = '"&de_emp_no&"' and de_id = '"&de_id&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into pay_yeartax_deduction (de_year,de_emp_no,de_id,de_emp_name,de_person_no,young_fdate,young_ldate,de_wonchen,de_tax_s,de_tax_w,de_tax_nation,de_tax_date,de_report_date,de_office,de_stay,de_position) values "
		sql = sql +	" ('"&de_year&"','"&de_emp_no&"','"&de_id&"','"&de_emp_name&"','"&de_person_no&"','"&young_fdate&"','"&young_ldate&"','"&de_wonchen&"','"&de_tax_s&"','"&de_tax_w&"','"&de_tax_nation&"','"&de_tax_date&"','"&de_report_date&"','"&de_office&"','"&de_stay&"','"&de_position&"')"
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
	response.write"location.replace('insa_pay_yeartax_deduction.asp');"	
	'response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
