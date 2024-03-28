<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	i_emp_no = request.form("i_emp_no")
	i_year = request.form("i_year")
	i_seq = request.form("i_seq")
	
	i_emp_name = request.form("i_emp_name")
	i_name = request.form("i_name")
	i_rel = request.form("i_rel")
	i_person_no = request.form("i_person_no")
	i_disab_chk = request.form("i_disab_chk")

	i_nts_amt =int(request.form("i_nts_amt"))
	i_other_amt =int(request.form("i_other_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_insurance set i_rel='"&i_rel&"',i_name='"&i_name&"',i_nts_amt='"&i_nts_amt&"',i_other_amt='"&i_other_amt&"' ,i_disab_chk='"&i_disab_chk&"' where i_year ='"&i_year&"' and i_emp_no = '"&i_emp_no&"' and i_person_no = '"&i_person_no&"' and i_seq = '"&i_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(i_seq) as max_seq from pay_yeartax_insurance where i_year='" + i_year + "' and i_emp_no='" + i_emp_no + "' and i_person_no='" + i_person_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			i_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			i_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_insurance (i_year,i_emp_no,i_person_no,i_seq,i_rel,i_name,i_nts_amt,i_other_amt,i_disab_chk) values "
		sql = sql +	" ('"&i_year&"','"&i_emp_no&"','"&i_person_no&"','"&i_seq&"','"&i_rel&"','"&i_name&"','"&i_nts_amt&"','"&i_other_amt&"','"&i_disab_chk&"')"
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
