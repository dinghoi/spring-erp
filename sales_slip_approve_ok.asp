<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	on Error resume next

	slip_id = request.form("slip_id")
	slip_no = request.form("slip_no")
	slip_seq = request.form("slip_seq")

	dbconn.BeginTrans

	Set xh = CreateObject("MSXML2.ServerXMLHTTP")
	   xh.open "GET","http://localhost/sales_slip_approve_html.asp?slip_id="&slip_id&"&slip_no="&slip_no&"&slip_seq="&slip_seq, false
	   xh.send()
	   strv = xh.ResponseBody
	Set xh = Nothing
	 
	fileName = ("d:\test.html")
	 
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso. FileExists(fileName) then fso.DeleteFile(fileName)
	Set fso = Nothing
	 
	Set f = CreateObject("ADODB.Stream")
	   f.open()
	   f.type = 1
	   f.write strv
	   f.savetofile fileName, 2
	   f.close
	Set f = Nothing

	sql = "Update sales_slip set sign_yn ='I' where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "전표 결재 요청중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "전표 결재 요청을 하였습니다."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"opener.document.frm.submit();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
