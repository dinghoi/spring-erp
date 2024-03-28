<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	acpt_no = int(request("acpt_no"))

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open DbConnect
	sql="delete from into where in_seq = 1 and acpt_no ="&acpt_no
	dbconn.execute(sql)

	sql = "Update as_acpt set as_process='접수',into_reason='',in_date=null,mod_date=getdate(),mod_id='"+user_id+"' where acpt_no = "&int(acpt_no)
	dbconn.execute(sql)

	response.write"<script language=javascript>"
	response.write"alert('입력이 완료되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>
	