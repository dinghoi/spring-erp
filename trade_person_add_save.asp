<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	trade_code = request.form("trade_code")
	person_name = request.form("person_name")
	person_grade = request.form("person_grade")
	person_tel_no = request.form("person_tel_no")
	person_email = request.form("person_email")
	person_memo = request.form("person_memo")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "Update trade_person set person_name='"&person_name&"',person_grade ='"&person_grade&"',person_tel_no='"&person_tel_no&"',person_email='"&person_email&"',person_memo ='"&person_memo&"',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where trade_code ='"&trade_code&"' and person_name ='"&person_name&"'"
		dbconn.execute(sql)
	  else
		sql = "select * from trade_person where trade_code ='"&trade_code&"' and person_name ='"&person_name&"'"
		Set rs=DbConn.Execute(Sql)
		if rs.eof or rs.bof then
			sql="insert into trade_person (trade_code,person_name,person_grade,person_tel_no,person_email,person_memo,reg_id,reg_name,reg_date) values ('"&trade_code&"','"&person_name&"','"&person_grade&"','"&person_tel_no&"','"&person_email&"','"&person_memo&"','"&user_id&"','"&user_name&"',now())"
			dbconn.execute(sql)
		  else
			response.write"<script language=javascript>"
			response.write"alert('이미 등록되어 있는 담당자입니다');"
			response.write"history.back();"
			response.write"</script>"
		end if
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "처리 되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
