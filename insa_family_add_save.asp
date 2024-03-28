<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	family_seq = request.form("family_seq")
	family_empno = request.form("family_empno")
	
	family_rel = request.form("family_rel")
	family_name = request.form("family_name")
	family_birthday = request.form("family_birthday")
	family_birthday_id = request.form("family_birthday_id")
	family_job = request.form("family_job")
	family_live = request.form("family_live")
	family_person1 = request.form("family_person1")
	family_person2 = request.form("family_person2")
	family_tel_ddd = request.form("family_tel_ddd")
    family_tel_no1 = request.form("family_tel_no1")
    family_tel_no2 = request.form("family_tel_no2")	
	family_support_yn = request.form("family_support_yn")
	if family_birthday = "" or isnull(family_birthday) then
	   family_birthday = "1900-01-01"
	end if
	family_national = request.form("family_national")
	family_witak = request.form("witak_check")
	family_holt = request.form("holt_check")
	family_holt_date = request.form("family_holt_date")
	family_pensioner = request.form("pensioner_check")
	family_serius = request.form("serius_check")
	family_merit = request.form("merit_check")
	family_disab = request.form("disab_check")
	family_children = request.form("children_check")
	if family_holt_date = "" then
	     family_holt_date = "1900-01-01"
	end if
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_family set family_rel='"&family_rel&"',family_name='"&family_name&"',family_birthday='"&family_birthday&"',family_birthday_id='"&family_birthday_id&"',family_job='"&family_job&"',family_live='"&family_live&"',family_support_yn='"&family_support_yn&"',family_person1='"&family_person1&"',family_person2='"&family_person2&"',family_tel_ddd='"&family_tel_ddd&"',family_tel_no1='"&family_tel_no1&"',family_tel_no2='"&family_tel_no2&"',family_national='"&family_national&"',family_witak='"&family_witak&"',family_holt='"&family_holt&"',family_holt_date='"&family_holt_date&"',family_pensioner='"&family_pensioner&"',family_serius='"&family_serius&"',family_merit='"&family_merit&"',family_disab='"&family_disab&"',family_children='"&family_children&"',family_mod_date= now(),family_mod_user='"&emp_user&"' where family_empno ='"&family_empno&"' and family_seq = '"&family_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(family_seq) as max_seq from emp_family where family_empno='" + family_empno + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			family_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			family_seq = right(max_seq,3)
		end if

		sql = "insert into emp_family (family_empno,family_seq,family_rel,family_name,family_birthday,family_birthday_id,family_job,family_live,family_support_yn,family_person1,family_person2,family_tel_ddd,family_tel_no1,family_tel_no2,family_national,family_disab,family_merit,family_serius,family_pensioner,family_witak,family_holt,family_holt_date,family_children,family_reg_date,family_reg_user) values "
		sql = sql +	" ('"&family_empno&"','"&family_seq&"','"&family_rel&"','"&family_name&"','"&family_birthday&"','"&family_birthday_id&"','"&family_job&"','"&family_live&"','"&family_support_yn&"','"&family_person1&"','"&family_person2&"','"&family_tel_ddd&"','"&family_tel_no1&"','"&family_tel_no2&"','"&family_national&"','"&family_disab&"','"&family_merit&"','"&family_serius&"','"&family_pensioner&"','"&family_witak&"','"&family_holt&"','"&family_holt_date&"','"&family_children&"',now(),'"&emp_user&"')"
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
