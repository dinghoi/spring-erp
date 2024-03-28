<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_bef = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

   
sql = "insert into emp_master_month select '201412' as emp_month,emp_master.* from emp_master"
'sql = "INSERT INTO emp_master_month ( emp_month ) SELECT emp_master.*, '201412' AS Expr1 FROM emp_master"
	response.write(sql)
	dbconn.execute(sql)
	   
dbconn.Close()
Set dbconn = Nothing
	
%>
