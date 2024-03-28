<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<%
'	on Error resume next
	user_id = request.form("user_id")
	emp_yn= request.form("emp_yn")
	pass = "1111"
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if emp_yn = "Y" then
		sql = "select * from emp_master where emp_no = '"+user_id+"'"
		set rs_emp=dbconn.execute(sql)
		pass = rs_emp("emp_person2")		
	end if

	sql = "Update memb set pass='"&pass&"',mod_id='"&mod_id&"',mod_date=now() where user_id = '"&user_id&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + " " + cstr(w_cnt) +" 건 등록 완료되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('초기화 되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
