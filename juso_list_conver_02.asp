<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Server.ScriptTimeOut = 10000  
' 대량 데이터 batch upload

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

Dbconn.BeginTrans

sql = "select * from juso_list where mg_ce_id > 'a' " 
Rs.Open Sql, Dbconn, 1 

i = 0
j = 0
do until rs.eof

	i = i + 1
	j = j + 1
	if j = 1000 then
		response.write("처리건수 : " + cstr(i))
		j = 0
	end if
	
	sql = "select user_id from memb where old_user_id = '"&rs("regi_id")&"'"
	set rs_memb=dbconn.execute(sql)
	if rs_memb.eof or rs_memb.bof then
		regi_id = ""
	  else
		regi_id = rs_memb("user_id")
	end if
	
	sql = "update juso_list set regi_id='"&regi_id&"' where company='"&rs("company")&"' and dept='"&rs("dept")&"'"
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