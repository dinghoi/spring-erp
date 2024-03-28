<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
pmg_yymm = request.form("pmg_yymm1")
etc_code = request.form("etc_code")

set dbconn = server.CreateObject("adodb.connection")
dbconn.open DbConnect

sql = "Update emp_etc_code set emp_payend_date='"&pmg_yymm&"',emp_payend_yn='Y' where emp_etc_code = '"&etc_code&"'"
dbconn.execute(sql)

end_msg = pmg_yymm & " 월 마감등록 되었습니다."

Response.write "<script type='text/javascript'>"
'response.write "	alert('마감등록 되었습니다....');"
Response.write "	alert('"&end_msg&"');"
Response.write "	parent.opener.location.reload();"
Response.write "	self.close() ;"
Response.write "</script>"

Response.End

dbconn.Close() : Set dbconn = Nothing

%>
