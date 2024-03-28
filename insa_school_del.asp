<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
position = request.cookies("nkpmg_user")("coo_position")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

sch_empno=Request.form("sch_empno")
sch_seq=Request.form("sch_seq")
sch_emp_name=Request.form("sch_emp_name")
owner_view=Request.form("owner_view")

'response.write(sch_empno)
'response.write(sch_seq)
'response.write(sch_emp_name)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_fam = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

    sql = " delete from emp_school " & _
	            "  where sch_empno ='"&sch_empno&"' and sch_seq = '"&sch_seq&"'"

	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "삭제중 Error가 발생하였습니다...."
	else
		' dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if
    'url = "insa_family_mg.asp?ck_sw="y"&view_condi=" + family_empno + "&ck_sw= y&view_condi="+view_condi+"&condi="+ condi
	'url = "insa_school_mg.asp?ck_sw=y&view_condi=" + sch_empno + "&condi="+ sch_emp_name
	if owner_view = "C" then
	           url = "insa_school_mg.asp?ck_sw=y&view_condi=" + sch_emp_name + "&owner_view="+ owner_view
		else
		       url = "insa_school_mg.asp?ck_sw=y&view_condi=" + sch_empno + "&owner_view="+ owner_view
	end if
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"location.replace('insa_family_mg.asp');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
