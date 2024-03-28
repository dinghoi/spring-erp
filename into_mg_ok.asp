<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	acpt_no = int(request("acpt_no"))
	in_seq = request.form("in_seq")
	if	in_seq = "" then
		in_seq = 0
	end if
	in_seq = int(in_seq) + 1
	in_process = request.form("in_process")
	into_date = request.form("into_date")
	in_place = request.form("in_place")
	in_remark = request.form("in_remark")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open DbConnect
	sql="insert into as_into (acpt_no,in_seq,in_process,into_date,in_place,in_remark,reg_id,reg_name,reg_date) values ('"&acpt_no&"','"&in_seq&"','"&in_process&"','"&into_date&"','"&in_place&"','"&in_remark&"','"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)

	if in_process = "대체회수" then
		in_process = "않함"
	end if	
	if in_process = "대체" or in_process = "않함" then
		sql = "Update as_acpt set in_replace ='"&in_process&"', mod_date=now(),mod_id='"&user_id&"' where acpt_no = "&int(acpt_no)
		dbconn.execute(sql)
	end if
		
	response.write"<script language=javascript>"
	response.write"alert('입력이 완료되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>
	