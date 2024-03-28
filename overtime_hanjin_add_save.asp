<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	work_date = request.form("work_date")
	mg_ce_id = request.form("mg_ce_id")
	company = request.form("company")
	dept = request.form("dept")
	from_hh = request.form("from_hh")
	from_mm = request.form("from_mm")
	from_time = cstr(from_hh) + cstr(from_mm)
	to_hh = request.form("to_hh")
	to_mm = request.form("to_mm")
	to_time = cstr(to_hh) + cstr(to_mm)
	work_gubun = request.form("work_gubun")
	work_item = request.form("work_gubun")
	work_memo = request.form("work_memo")
	you_yn = request.form("you_yn")
	sign_no = request.form("sign_no")
	cancel_yn = request.form("cancel_yn")
	acpt_no = 0

	dbconn.BeginTrans

	sql = "select * from overtime_code where work_gubun = '"&work_gubun&"'"
	set rs_etc=dbconn.execute(sql)
	cost_detail = rs_etc("cost_detail")
	overtime_amt = rs_etc("overtime_amt")

	sql = "delete from overtime where work_date ='"&work_date&"' and mg_ce_id='"&mg_ce_id&"'"
	dbconn.execute(sql)

	sql = "select * from memb where user_id = '"&mg_ce_id&"'"
	set rs_memb=dbconn.execute(sql)

	sql="insert into overtime (work_date,mg_ce_id,user_name,user_grade,acpt_no,emp_company,bonbu,saupbu,team,org_name,reside_place,company,dept,work_item"& _
	",from_time,to_time,work_gubun,cost_detail,person_amt,overtime_amt,work_memo,sign_no,you_yn,cancel_yn,end_yn,reg_id,reg_user,reg_date)"& _
	" values ('"&work_date&"','"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"',"&acpt_no&",'"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&company&"','"&dept&"','"&work_item&"','"&from_time&"','"&to_time&"','"&work_gubun&"','"&cost_detail&"','1',"&overtime_amt&",'"&work_memo&"','"&sign_no&"','"&you_yn&"','"&cancel_yn&"','N','"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "이미 등록되어 있어 등록중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>
