<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	cost_month = request.form("cost_month")
	saupbu = request.form("saupbu")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if emp_no = "100001" then
		sql = "Update cost_end set ceo_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
	  else
		sql = "Update cost_end set bonbu_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
	end if
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "승인중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "승인되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"parent.opener.location.reload();"
	response.write"opener.document.frm.submit();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
