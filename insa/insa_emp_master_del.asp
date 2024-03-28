<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

'user_name = request.cookies("nkpmg_user")("coo_user_name")
'user_id = request.cookies("nkpmg_user")("coo_user_id")
'position = request.cookies("nkpmg_user")("coo_position")
'insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

emp_no=Request.form("emp_no")
emp_name=Request.form("emp_name")
emp_company=Request.form("emp_company")
view_condi=Request.form("view_condi")

'response.write(emp_no)
'response.write(emp_name)
'response.write(emp_company)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_fam = Server.CreateObject("ADODB.Recordset")
Set rs_sch = Server.CreateObject("ADODB.Recordset")
Set rs_car = Server.CreateObject("ADODB.Recordset")
Set rs_qual = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_lang = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_acpt = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql="select * from emp_master where emp_no = '"&emp_no&"'"
set Rs_emp=dbconn.execute(sql)
if not Rs_emp.eof then

    sql = " delete from emp_master " & _
	            "  where emp_no = '"&emp_no&"'"
	dbconn.execute(sql)

	sql="select * from emp_family where family_empno = '"&emp_no&"'"
    set Rs_fam=dbconn.execute(sql)
    if not Rs_fam.eof then
           sql = " delete from emp_family " & _
	             "  where family_empno = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	sql="select * from emp_school where sch_empno = '"&emp_no&"'"
    set rs_sch=dbconn.execute(sql)
    if not rs_sch.eof then
           sql = " delete from emp_school " & _
	             "  where sch_empno = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	sql="select * from emp_career where career_empno = '"&emp_no&"'"
    set rs_car=dbconn.execute(sql)
    if not rs_car.eof then
           sql = " delete from emp_career " & _
	             "  where career_empno = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	sql="select * from emp_qual where qual_empno = '"&emp_no&"'"
    set rs_qual=dbconn.execute(sql)
    if not rs_qual.eof then
           sql = " delete from emp_qual " & _
	             "  where qual_empno = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	sql="select * from emp_edu where edu_empno = '"&emp_no&"'"
    set rs_edu=dbconn.execute(sql)
    if not rs_edu.eof then
           sql = " delete from emp_edu " & _
	             "  where edu_empno = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	sql="select * from emp_language where lang_empno = '"&emp_no&"'"
    set rs_lang=dbconn.execute(sql)
    if not rs_lang.eof then
           sql = " delete from emp_language " & _
	             "  where lang_empno = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	sql="select * from memb where user_id = '"&emp_no&"'"
	set rs_memb=dbconn.execute(sql)
    if not rs_memb.eof then
           sql = " delete from memb " & _
	             "  where user_id = '"&emp_no&"'"
	       dbconn.execute(sql)
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "삭제중 Error가 발생하였습니다...."
	else
'		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if

end if

if view_condi = "성명" then
		   url = "insa_master_modify.asp?ck_sw=y&condi=" + emp_name + "&view_condi="+ view_condi
	else
		   url = "insa_master_modify.asp?ck_sw=y&condi=" + emp_no + "&view_condi="+ view_condi
end If

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
'Response.write "	location.replace('insa_master_modify.asp');"
Response.write "	location.replace('"&url&"');"
Response.write "</script>"
Response.End

dbconn.Close() : Set dbconn = Nothing
%>
