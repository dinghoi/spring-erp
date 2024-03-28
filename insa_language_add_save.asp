<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	lang_seq = request.form("lang_seq")
	lang_empno = request.form("lang_empno")
	
	lang_id = request.form("lang_id")
    lang_id_type = request.form("lang_id_type")
    lang_point = request.form("lang_point")
    lang_grade = request.form("lang_grade")
    lang_get_date = request.form("lang_get_date")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_language set lang_id='"&lang_id&"',lang_id_type='"&lang_id_type&"',lang_point='"&lang_point&"',lang_grade='"&lang_grade&"',lang_get_date='"&lang_get_date&"',lang_mod_date=now(),lang_mod_user='"&emp_user&"' where lang_empno ='"&lang_empno&"' and lang_seq = '"&lang_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(lang_seq) as max_seq from emp_language where lang_empno='" + lang_empno + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			lang_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			lang_seq = right(max_seq,3)
		end if

		sql = "insert into emp_language (lang_empno,lang_seq,lang_id,lang_id_type,lang_point,lang_grade,lang_get_date,lang_reg_date,lang_reg_user) values "
		sql = sql +	" ('"&lang_empno&"','"&lang_seq&"','"&lang_id&"','"&lang_id_type&"','"&lang_point&"','"&lang_grade&"','"&lang_get_date&"',now(),'"&emp_user&"')"
		
		'response.write sql
		
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
