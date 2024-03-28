<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = now()
be_date = dateadd("m",-1,curr_date)
be_month = cstr(mid(be_date,1,4)) + cstr(mid(be_date,6,2))

'response.write(emp_yymm)
'response.write(view_condi)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_bef = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Dbconn.BeginTrans 

sql = "delete from emp_org_mst_month where org_month ='"&be_month&"'"
    dbconn.execute(sql)
	
sql = "delete from emp_master_month where emp_month ='"&be_month&"'"
    dbconn.execute(sql)	

sql = "insert into emp_org_mst_month select '"&be_month&"' as org_month,emp_org_mst.* from emp_org_mst"
    dbconn.execute(sql)

sql = "insert into emp_master_month select '"&be_month&"' as emp_month,emp_master.* from emp_master"
	dbconn.execute(sql)
	
end_sw = "Y" 

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	response.write"<script language=javascript>"
	response.write"alert('"&be_month&"...조직 및 인사 마스타 마감처리가 되었습니다...');"		
	'response.write"location.replace('insa_master_month_mg.asp');"
	response.write"location.replace('insa_org_mg.asp');"
	response.write"</script>"
	Response.End
end if

dbconn.Close()
Set dbconn = Nothing
	
%>
