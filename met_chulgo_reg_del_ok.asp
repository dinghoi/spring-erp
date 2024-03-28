<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	old_rele_no = abc("old_rele_no")
	old_rele_seq = abc("old_rele_seq")
	old_rele_date = abc("old_rele_date")
	old_att_file = abc("old_att_file")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "delete from met_chulgo_reg where rele_no ='"&old_rele_no&"' and rele_seq='"&old_rele_seq&"' and rele_date='"&old_rele_date&"'"
	dbconn.execute(sql)
	sql = "delete from met_chulgo_reg_goods where rele_no ='"&old_rele_no&"' and rele_seq='"&old_rele_seq&"' and rele_date='"&old_rele_date&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
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
