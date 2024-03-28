<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	on Error resume next

	u_type = request.form("u_type")
	company = request.form("company")
	high_org = request.form("high_org")
	org_first = request.form("org_first")
	org_second = request.form("org_second")
	dept_name = request.form("dept_name")
	person = request.form("person")
	sido = request.form("sido")
	gugun = request.form("gugun")
	dong = request.form("dong")
	addr = request.form("addr")
	tel_ddd = request.form("tel_ddd")
	tel_no1 = request.form("tel_no1")
	tel_no2 = request.form("tel_no2")
	internet_no = request.form("internet_no")
	internet_pass = request.form("internet_pass")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	if u_type = "U" then
		dept_code = request.form("dept_code")
		sql = "update asset_dept set high_org='"&high_org&"', org_first='"&org_first&"', org_second='"&org_second&"', dept_name='"&dept_name&"', person='"&person&"', person='"&person&"', tel_ddd='"&tel_ddd&"', tel_no1='"&tel_no1&"', tel_no2='"&tel_no2&"', sido='"&sido&"', gugun='"&gugun&"', dong='"&dong&"', addr='"&addr&"', internet_no='"&internet_no&"', mod_id='"&reg_id&"', mod_date=now() where company='" + company + "' and dept_code = '" + dept_code + "'"
		dbconn.execute(sql)	  
	  else
		sql="select max(dept_code) as max_seq from asset_dept where company='" + company + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			dept_code = "1001"
		  else
			dept_code = cstr((int(rs("max_seq")) + 1))
		end if
	
		sql="insert into asset_dept (company,dept_code,high_org,org_first,org_second,dept_name,person,tel_ddd,tel_no1,tel_no2,sido,gugun,dong,addr,internet_no,reg_id,reg_date) values ('"&company&"','"&dept_code&"','"&high_org&"','"&org_first&"','"&org_second&"','"&dept_name&"','"&person&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&internet_no&"','"&reg_id&"',now())"
		dbconn.execute(sql)
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록 완료되었습니다...."
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
