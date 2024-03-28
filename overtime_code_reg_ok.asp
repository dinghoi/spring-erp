<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	u_type = request.form("u_type")
	overtime_code = request.form("overtime_code")
	work_gubun = request.form("work_gubun")
	holi_id = request.form("holi_id")
	cost_detail = holi_id + "수당"
	apply_dept = request.form("apply_dept")
	apply_unit = request.form("apply_unit")
	overtime_amt = int(request.form("overtime_amt"))
	meals_yn = request.form("meals_yn")
	work_time1 = request.form("work_time1")
	work_time2 = int(request.form("work_time2"))
	sign_yn = request.form("sign_yn")
	you_yn = request.form("you_yn")
	overtime_memo = request.form("overtime_memo")
	use_yn = request.form("use_yn")

	if u_type <> "U" then
		Sql="select max(overtime_code) from overtime_code"
		Set rs=DbConn.Execute(Sql)
		last_no = cint(rs(0))
		last_no = last_no + 1
		overtime_code = right("0" + cstr(last_no),2)
	end if

	sql = "delete from overtime_code where overtime_code ='"&overtime_code&"'"
	dbconn.execute(sql)

	sql="insert into overtime_code  values ('"&overtime_code&"','"&work_gubun&"','"&cost_detail&"','"&holi_id&"','"&apply_dept&"','"&apply_unit&"',"&overtime_amt&",'"&meals_yn&"','"&work_time1&"',"&work_time2&",'"&sign_yn&"','"&you_yn&"','"&overtime_memo&"','"&use_yn&"','"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.Redirect "overtime_code_mg.asp"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
