<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	trade_code = request.form("trade_code")
	trade_no1 = request.form("trade_no1")
	trade_no2 = request.form("trade_no2")
	trade_no3 = request.form("trade_no3")
	trade_no = cstr(trade_no1) + cstr(trade_no2) + cstr(trade_no3)
	old_trade_no = request.form("old_trade_no")
	tradename = request.form("trade_name")
	trade_name = replace(tradename,"（주）","(주)")
	trade_id = request.form("trade_id")
	sales_type = request.form("sales_type")
	trade_owner = request.form("trade_owner")
	trade_addr = request.form("trade_addr")	
	trade_uptae = request.form("trade_uptae")	
	trade_upjong = request.form("trade_upjong")
	trade_tel = request.form("trade_tel")	
	trade_fax = request.form("trade_fax")
	group_name = request.form("group_name")
	use_sw = "Y"
	person_name = request.form("person_name")
	person_grade = request.form("person_grade")
	person_tel_no = request.form("person_tel_no")
	person_email = request.form("person_email")
	person_memo = request.form("person_memo")
	emp_no = request.form("emp_no")
	emp_name = request.form("emp_name")
	saupbu = request.form("saupbu")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select max(trade_code) as max_seq from trade"
	set rs=dbconn.execute(sql)
			
	if	isnull(rs("max_seq"))  then
		trade_code = "00001"
	  else
		max_seq = "0000" + cstr((int(rs("max_seq")) + 1))
		trade_code = right(max_seq,5)
	end if

	sql = "select * from trade where trade_no ='"&trade_no&"' AND (trade_name ='"&trade_name&"' or trade_full_name ='"&trade_full_name&"')"
	Set rs=DbConn.Execute(Sql)
	'Response.write Sql
	if rs.eof or rs.bof then
		sql="insert into trade (trade_code,trade_no,trade_name,trade_id,sales_type,trade_owner,trade_addr,trade_uptae,trade_upjong,trade_tel,trade_fax,mg_group,group_name,emp_no,emp_name,saupbu,use_sw,reg_id,reg_date) values ('"&trade_code&"','"&trade_no&"','"&trade_name&"','"&trade_id&"','"&sales_type&"','"&trade_owner&"','"&trade_addr&"','"&trade_uptae&"','"&trade_upjong&"','"&trade_tel&"','"&trade_fax&"','"&mg_group&"','"&group_name&"','"&emp_no&"','"&emp_name&"','"&saupbu&"','"&use_sw&"','"&user_id&"',now())"
		dbconn.execute(sql)
		'Response.write sql
		if (person_name <> "" or isnull(person_name)) and (person_email <> "" or isnull(person_name)) then
			sql="insert into trade_person (trade_code,person_name,person_grade,person_tel_no,person_email,person_memo,reg_id,reg_name,reg_date) values ('"&trade_code&"','"&person_name&"','"&person_grade&"','"&person_tel_no&"','"&person_email&"','"&person_memo&"','"&user_id&"','"&user_name&"',now())"
			dbconn.execute(sql)
			'Response.write sql
		end if
	else
		response.write"<script language=javascript>"
		response.write"alert('이미 등록되어 있는 거래처입니다');"
		response.write"history.back();"
		response.write"</script>"
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
	response.write"opener.document.frm.submit();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
