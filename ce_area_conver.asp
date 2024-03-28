<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
' 대량 데이터 batch upload

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

Dbconn.BeginTrans

sql = "select * from ce_area" 
Rs.Open Sql, Dbconn, 1 

i = 0
do until rs.eof

	i = i + 1
	
	sql = "select user_id from memb where old_user_id = '"&rs("mg_ce_id")&"'"
	set rs_memb=dbconn.execute(sql)
	mg_ce_id = rs_memb("user_id")

	sql = "select user_id from memb where old_user_id = '"&rs("mod_id")&"'"
	set rs_memb=dbconn.execute(sql)
	if rs_memb.eof or rs_memb.bof then
		mod_id = ""
	  else
		mod_id = rs_memb("user_id")
	end if
	
	sql = "update ce_area set mg_ce_id='"&mg_ce_id&"', mod_id='"&mod_id&"' where sido='"&rs("sido")&"' and gugun = '"&rs("gugun")&"' and mg_ce_id = '"&rs("mg_ce_id")&"'"
	dbconn.execute(sql)	  

	rs.movenext()
loop

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
end if

response.write("처리건수 : " + cstr(i))

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
%>