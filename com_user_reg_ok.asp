<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	new_user_id = request.form("user_id")
	old_user_id = request.form("old_user_id")
	company = request.form("company")
	pass = request.form("pass")
	reside = request.form("reside")
	grade = request.form("grade")

	dbconn.BeginTrans

	if u_type = "U" then
		sql = "Update memb set user_name='"&company&"',pass='"&pass&"',grade='"&grade&"',reside='"&reside&"',mod_id='"&user_id&"',mod_date=now() where user_id = '"&old_user_id&"'"
		dbconn.execute(sql)
	  else
		sql="insert into memb (user_id,pass,emp_no,user_name,user_grade,team,org_name,hp,grade,mg_group,reside,sms,help_yn,cost_grade,pay_grade,insa_grade,account_grade,sales_grade,reg_id,reg_name,reg_date,login_cnt,login_date) values ('"&new_user_id&"','"&pass&"','999999','"&company&"','회사','사용자','사용자','.','"&grade&"','1','"&reside&"','N','N','5','3','3','2','3','"&user_id&"','"&user_name&"',now(),'0',now())"
		dbconn.execute(sql)
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "이미 등록되어 있어 등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"		
	response.Redirect "com_user_mg.asp"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
