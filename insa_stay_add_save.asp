<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	stay_code = request.form("stay_code")
	
	'response.write family_empno
	'response.write"alert('"&family_empno&"');"
	
    stay_name = request.form("stay_name")
	stay_org_code = request.form("stay_org_code")
	stay_org_name = request.form("stay_org_name")
	stay_reside_company = request.form("stay_reside_company")
    stay_sido = request.form("stay_sido")
    stay_gugun = request.form("stay_gugun")
    stay_dong = request.form("stay_dong")
    stay_addr = request.form("stay_addr")
    stay_tel_ddd = request.form("stay_tel_ddd")
    stay_tel_no1 = request.form("stay_tel_no1")
    stay_tel_no2 = request.form("stay_tel_no2")
    stay_fax_ddd = request.form("stay_fax_ddd")
    stay_fax_no1 = request.form("stay_fax_no1")
    stay_fax_no2 = request.form("stay_fax_no2")
    stay_reg_date = now()

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans


	if	u_type = "U" then
		sql = "update emp_stay set stay_name='"&stay_name&"',set stay_org_code='"&stay_org_code&"',set stay_org_name='"&stay_org_name&"',set stay_reside_company='"&stay_reside_company&"',stay_sido='"&stay_sido&"',stay_gugun='"&stay_gugun&"',stay_dong='"&stay_dong&"',stay_addr='"&stay_addr&"',stay_tel_ddd='"&stay_tel_ddd&"',stay_tel_no1='"&stay_tel_no1&"',stay_tel_no2='"&stay_tel_no2&"',stay_fax_ddd='"&stay_fax_ddd&"',stay_fax_no1='"&stay_fax_no1&"',stay_fax_no2='"&stay_fax_no2&"' where stay_code ='"&stay_code&"'"
		dbconn.execute(sql)	  
	  else
		sql = "insert into emp_stay(stay_code,stay_name,stay_org_code,stay_org_name,stay_reside_company,stay_sido,stay_gugun,stay_dong,stay_addr,stay_tel_ddd,stay_tel_no1,stay_tel_no2,stay_fax_ddd,stay_fax_no1,stay_fax_no2,stay_reg_date) values "
		sql = sql +	" ('"&stay_code&"','"&stay_name&"','"&stay_org_code&"','"&stay_org_name&"','"&stay_reside_company&"','"&stay_sido&"','"&stay_gugun&"','"&stay_dong&"','"&stay_addr&"','"&stay_tel_ddd&"','"&stay_tel_no1&"','"&stay_tel_no2&"','"&stay_fax_ddd&"','"&stay_fax_no1&"','"&stay_fax_no2&"',now())"
		dbconn.execute(sql)
	end if

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
