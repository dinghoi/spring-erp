<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	app_empno = request.form("app_empno")
	app_emp_name = request.form("app_emp_name")
	
	apphu_seq = request.form("apphu_seq")
    apphu_id_type = request.form("apphu_id_type")
	apphu_date = request.form("apphu_date")
    apphu_start_date = request.form("apphu_start_date")
    apphu_finish_date = request.form("apphu_finish_date")
	
	app_to_company = request.form("app_company")
	app_to_orgcode = request.form("app_org")
	app_to_org = request.form("app_org_name")
	app_to_grade = request.form("app_grade")
	app_to_job = request.form("app_job")
	app_to_position = request.form("app_position")
	
    app_date = request.form("app_bok_date")
    app_seq = request.form("app_bok_seq")
    app_comment = request.form("app_comment")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
    Set Rs_app = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

'    휴직발령.. 복직으로 셋팅 update를 위한...
		sql = "update emp_appoint set app_bokjik_id ='Y' where app_empno = '"&app_empno&"' and app_seq = '"&apphu_seq&"' and app_id = '휴직발령' and app_date = '"&apphu_date&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
		
'    복직발령 등록......

		sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_start_date,app_finish_date,app_comment,app_reg_date) values "
		sql = sql +	" ('"&app_empno&"','"&app_seq&"','복직발령','"&app_date&"','"&app_emp_name&"','"&apphu_id_type&"','"&app_to_company&"','"&app_orgcode&"','"&app_org&"','"&app_to_grade&"','"&app_to_job&"','"&app_to_position&"','"&apphu_start_date&"','"&apphu_finish_date&"','"&app_comment&"',now())"
		
		'response.write sql		
		
		dbconn.execute(sql)

'    인사마스터에 급여대상 "0"으로 풀어주기.....

        sql = "update emp_master set emp_pay_id ='0' where emp_no = '"&app_empno&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  


	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
