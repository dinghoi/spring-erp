<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	e_emp_no = request.form("e_emp_no")
	e_year = request.form("e_year")
	e_seq = request.form("e_seq")
	
	e_emp_name = request.form("e_emp_name")
	e_name = request.form("e_name")
	e_rel = request.form("e_rel")
	e_person_no = request.form("e_person_no")
	e_edu_level = request.form("e_edu_level")
	e_disab = request.form("e_disab")
	e_uniform = request.form("e_uniform")

	e_nts_amt =int(request.form("e_nts_amt"))
	e_other_amt =int(request.form("e_other_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_edu set e_rel='"&e_rel&"',e_name='"&e_name&"',e_disab='"&e_disab&"',e_uniform='"&e_uniform&"',e_edu_level='"&e_edu_level&"',e_nts_amt='"&e_nts_amt&"',e_other_amt='"&e_other_amt&"' where e_year ='"&e_year&"' and e_emp_no = '"&e_emp_no&"' and e_person_no = '"&e_person_no&"' and e_seq = '"&e_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(e_seq) as max_seq from pay_yeartax_edu where e_year='" + e_year + "' and e_emp_no='" + e_emp_no + "' and e_person_no='" + e_person_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			e_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			e_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_edu (e_year,e_emp_no,e_person_no,e_seq,e_rel,e_name,e_disab,e_edu_level,e_uniform,e_nts_amt,e_other_amt) values "
		sql = sql +	" ('"&e_year&"','"&e_emp_no&"','"&e_person_no&"','"&e_seq&"','"&e_rel&"','"&e_name&"','"&e_disab&"','"&e_edu_level&"','"&e_uniform&"','"&e_nts_amt&"','"&e_other_amt&"')"
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
