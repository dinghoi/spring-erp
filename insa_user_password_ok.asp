<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	view_condi = request.form("view_condi1")

	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	Sql = "SELECT * FROM emp_master where emp_no = '"&view_condi&"'"
    Set rs_emp = DbConn.Execute(SQL)
	if not rs_emp.eof then
	       emp_person2 = rs_emp("emp_person2")
		   if emp_person2 = "" or isnull(emp_person2) then
			     emp_person2 = view_condi
		   end if
    end if
	
	sql = "Update memb set pass='"&emp_person2&"',mod_id ='"&user_id&"',mod_date=now() where user_id='"&view_condi&"'"
	dbconn.execute(sql)
	
	response.write"<script language=javascript>"
	response.write"alert('변경되었습니다....');"		
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	
	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>
	