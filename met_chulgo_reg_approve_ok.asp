<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	on Error resume next

	rele_no = request.form("rele_no")
	rele_seq = request.form("rele_seq")
	rele_date = request.form("rele_date")

	dbconn.BeginTrans

	Set xh = CreateObject("MSXML2.ServerXMLHTTP")
	   xh.open "GET","http://localhost/met_chulgo_reg_approve_html.asp?rele_no="&rele_no&"&rele_seq="&rele_seq&"&rele_date="&rele_date, false
	   xh.send()
	   strv = xh.ResponseBody
	Set xh = Nothing
	 
	fileName = ("d:\met_chulgo_reg_test.html")
	 
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

'	sql = "Update met_chulgo_reg set rele_sign_yn ='I' where rele_no = '"&rele_no&"' and rele_seq = '"&rele_seq&"' and rele_date = '"&rele_date&"'"   테스트를 하기위해 강제결재 완료처리 - 추후 그룹웨어 결재 연계
	sql = "Update met_chulgo_reg set rele_sign_yn ='Y' where rele_no = '"&rele_no&"' and rele_seq = '"&rele_seq&"' and rele_date = '"&rele_date&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "출고의뢰 결재 요청중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "출고의뢰 결재 요청을 하였습니다."
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
