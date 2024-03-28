<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)


'sql = "update emp_master_month set cost_group='한진그룹' where cost_group = '한진'"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Dbconn.BeginTrans

i = 0

sql = "update pay_month_give set cost_group='한진그룹' where cost_group = '한진'"

dbconn.execute(sql)

sql = "update pay_month_deduct set cost_group='한진그룹' where cost_group = '한진'"

dbconn.execute(sql)


if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	response.write"<script language=javascript>"
	response.write"alert('"&i&"...한진처리 되었습니다...');"		
	'response.write"location.replace('insa_master_month_mg.asp');"
	response.write"location.replace('insa_person_mg.asp');"
	response.write"</script>"
	Response.End
end if

dbconn.Close()
Set dbconn = Nothing
	
%>
