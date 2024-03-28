<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	o_year = request.form("inc_yyyy")
	o_emp_no = request.form("emp_no")
	o_emp_name = request.form("emp_name")
	o_person_no = request.form("emp_person")

	'response.write(y_emp_no)
	'response.End

	o_nps =int(request.form("o_nps"))
	o_nhis =int(request.form("o_nhis"))
	o_sosang =int(request.form("o_sosang"))
	o_chul2012 =int(request.form("o_chul2012"))
	o_chul2013 =int(request.form("o_chul2013"))
	o_chul2014 =int(request.form("o_chul2014"))
	o_woori =int(request.form("o_woori"))
	o_goyoung =int(request.form("o_goyoung"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_other set o_nps='"&o_nps&"',o_nhis='"&o_nhis&"',o_sosang='"&o_sosang&"',o_chul2012='"&o_chul2012&"',o_chul2013='"&o_chul2013&"',o_chul2014='"&o_chul2014&"',o_woori='"&o_woori&"',o_goyoung='"&o_goyoung&"' where o_year ='"&o_year&"' and o_emp_no = '"&o_emp_no&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into pay_yeartax_other (o_year,o_emp_no,o_emp_name,o_person_no,o_nps,o_nhis,o_sosang,o_chul2012,o_chul2013,o_chul2014,o_woori,o_goyoung) values "
		sql = sql +	" ('"&o_year&"','"&o_emp_no&"','"&o_emp_name&"','"&o_person_no&"','"&o_nps&"','"&o_nhis&"','"&o_sosang&"','"&o_chul2012&"','"&o_chul2013&"','"&o_chul2014&"','"&o_woori&"','"&o_goyoung&"')"
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
	response.write"location.replace('insa_pay_yeartax_other.asp');"	
	'response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
