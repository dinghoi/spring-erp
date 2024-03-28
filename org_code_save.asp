<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	company = request.form("company")
	org_gubun = request.form("org_gubun")
	org_name = request.form("org_name")
	org_name = replace(org_name," ","")
	used_sw = request.form("used_sw")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if u_type = "U" then
		org_code = request.form("org_code")
		sql = "update org_code set org_name='"+org_name+"', used_sw='"+used_sw+"', reg_id='"+user_id+"', reg_date=now() where org_company='"+company+"' and org_gubun = '"+org_gubun+"' and org_code='"+org_code+"'"
	  else
		sql="select max(org_code) as max_seq from org_code where org_company='" + company + "' and org_gubun = '" + org_gubun + "'"
		set rs=dbconn.execute(sql)
			
		if	isnull(rs("max_seq"))  then
			code_seq = "01"
		  else
			max_seq = "0" + cstr((int(rs("max_seq")) + 1))
			code_seq = right(max_seq,2)
		end if
	
		sql = "insert into org_code (org_company,org_gubun,org_code,org_name,used_sw,reg_id,reg_date) values ('"&company&"','"&org_gubun&"','"&code_seq&"','"&org_name&"','"&used_sw&"','"&user_id&"',now())"
	end if
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
