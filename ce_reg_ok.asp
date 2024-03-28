<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<%
	u_type = request.form("u_type")
	out_yn = request.form("out_yn")
	user_id = request.form("user_id")
	pass = "1111"
	user_name = request.form("user_name")
	user_grade = request.form("user_grade")
	user_type = request.form("user_type")
	if out_yn = "Y" then
		if user_type = "Y" then
			emp_company = request.form("emp_company")
			bonbu = request.form("bonbu")
			saupbu = request.form("saupbu")
			team = request.form("team")
			org_name = request.form("org_name")
			reside_place = request.form("reside_place")
			reside_company = request.form("reside_company")
		  else
			emp_company = ""
			bonbu = ""
			saupbu = ""
			team = request.form("team")
			org_name = request.form("org_name")
			reside_place = request.form("reside_place1")
			reside_company = request.form("reside_company1")
		end if
	end if
	hp = request.form("hp")
	email = request.form("email")
	grade = request.form("grade")
	help_yn = request.form("help_yn")
	mg_group = request.form("mg_group")
	if isnull(reside_place) or reside_place = "" then
		reside = "0"
	  else
	  	reside = "1"
	end if

	sms = request.form("sms")
	
	mod_id = request.Cookies("nkpmg_user")("coo_user_id")
	mod_name = request.Cookies("nkpmg_user")("coo_user_name")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	if	u_type = "U" then
		if out_yn = "Y" then
			sql = "Update memb set user_name='"&user_name&"',user_grade='"&user_grade&"',emp_company ='"&emp_company&"',bonbu ='"&bonbu&"',saupbu ='"&saupbu&"',team ='"&team&"',org_name ='"&org_name&"',reside_place='"&reside_place&"',reside_company ='"&reside_company&"',reside='"&reside&"',hp='"&hp&"',email='"&email&"',grade='"&grade&"',help_yn='"+help_yn+"',mg_group='"&mg_group&"',sms='"&sms&"',mod_id='"&mod_id&"',mod_date=now() where user_id = '"&user_id&"'"
		  else
			sql = "Update memb set hp='"&hp&"',email='"&email&"',grade='"&grade&"',help_yn='"+help_yn+"',mg_group='"&mg_group&"',sms='"&sms&"',mod_id='"&mod_id&"',mod_date=now() where user_id = '"&user_id&"'"
		end if
		dbconn.execute(sql)
	  else
		sql="insert into memb (user_id,pass,emp_no,user_name,user_grade,emp_company,bonbu,saupbu,team,org_name,reside_place,reside_company,hp,email,grade,mg_group,reside,sms,help_yn,cost_grade,pay_grade,insa_grade,account_grade,sales_grade,reg_id,reg_name,reg_date) values ('"&user_id&"','1111','999999','"&user_name&"','"&user_grade&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&reside_company&"','"&hp&"','"&email&"','"&grade&"','"&mg_group&"','"&reside&"','"&sms&"','N','5','3','3','3','3','"&mod_id&"','"&mod_name&"',now())"
		dbconn.execute(sql)
	end if	

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
