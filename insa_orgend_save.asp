<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next
	
    curr_date = mid(cstr(now()),1,10)
	
	u_type = request.form("u_type")
	mod_user = request.cookies("nkpmg_user")("coo_user_name")
	
	org_code = request.form("org_code")
	org_end_date = request.form("org_end_date")

    'response.write(org_code)
	'response.write(org_end_date)
    
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "update emp_org_mst set org_end_date='"&org_end_date&"' where org_code = '"&org_code&"'"

		'response.write sql
		
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
