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

	old_buy_no = abc("old_buy_no")
	old_buy_date = abc("old_buy_date")
	old_buy_goods_type = abc("old_buy_goods_type")
	old_att_file = abc("old_att_file")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "delete from met_buy where buy_no ='"&old_buy_no&"' and buy_date='"&old_buy_date&"' and buy_goods_type='"&old_buy_goods_type&"'"
	dbconn.execute(sql)
	sql = "delete from met_buy_goods where bg_no ='"&old_buy_no&"' and bg_date='"&old_buy_date&"' and bg_goods_type='"&old_buy_goods_type&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "������ Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = "�����Ǿ����ϴ�...."
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
