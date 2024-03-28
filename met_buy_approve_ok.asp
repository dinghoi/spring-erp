<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	on Error resume next

	buy_no = request.form("buy_no")
	buy_date = request.form("buy_date")
	buy_seq = request.form("buy_seq")

	dbconn.BeginTrans

	Set xh = CreateObject("MSXML2.ServerXMLHTTP")
	   xh.open "GET","http://localhost/met_buy_approve_html.asp?buy_no="&buy_no&"&buy_date="&buy_date&"&buy_seq="&buy_seq, false
	   xh.send()
	   strv = xh.ResponseBody
	Set xh = Nothing
	 
	fileName = ("d:\met_buy_test.html")
	 
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

'	sql = "Update met_buy set buy_sign_yn ='I' where buy_no = '"&buy_no&"' and buy_date = '"&buy_date&"' and buy_seq = '"&buy_seq&"'"
' 테스트를 하기위한 부분으로 그룹웨어 결재가 정상적으로 되면 위 sql로 교체	
	sql = "Update met_buy set buy_sign_yn ='Y' where buy_no = '"&buy_no&"' and buy_date = '"&buy_date&"' and buy_seq = '"&buy_seq&"'"
	
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "구매품의 결재 요청중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "구매품의 결재 요청을 하였습니다."
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
