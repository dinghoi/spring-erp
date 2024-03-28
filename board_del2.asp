<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	pass = request.form("pass")
	board_seq = request.form("board_seq")
	page = request("page")
	condi = request("condi")
	condi_value = request("condi_value")
	ck_sw = Request("ck_sw")
	
	Set dbconn = server.CreateObject("adodb.connection")
	Set Rs1 = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	Sql="select * from board2 where board_seq="&board_seq
	Set Rs=dbconn.execute(Sql)

	if	rs("pass") <> pass then
		response.write"<script language=javascript>"
		response.write"alert('입력하신 비밀번호가 틀립니다.');"
		response.write"history.go(-1);"
		response.write"</script>"
	  Else
		sql="delete from board2 where board_seq="&board_seq
		dbconn.execute(sql)
		url = "nkp_main2.asp?page="&page&"&condi="&condi&"&condi_value="&condi_value&"&ck_sw=y"
		response.write"<script language=javascript>"
		response.write"alert('삭제 되었습니다.');"
		response.write"location.replace('"&url&"');"
		response.write"</script>"		
	End If

	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>
