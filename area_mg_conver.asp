<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Server.ScriptTimeOut = 100000  
' �뷮 ������ batch upload

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

Dbconn.BeginTrans

sql = "select * from area_mg where mg_group ='1' and mg_ce_id > 'a'" 
Rs.Open Sql, Dbconn, 1 

i = 0
j = 0
do until rs.eof

	i = i + 1
	j = j + 1
	if j = 1000 then
		response.write("ó���Ǽ� : " + cstr(i))
		j = 0
	end if
	
	sql = "select mg_ce_id from ce_area where sido = '"&rs("sido")&"' and gugun = '"&rs("gugun")&"' and mg_group = '"&rs("mg_group")&"' "
	set rs_ce=dbconn.execute(sql)
	if rs_ce.eof or rs_ce.bof then
		mg_ce_id = "error"
	  else
		mg_ce_id = rs_ce("mg_ce_id")
	end if
	
	sql = "update area_mg set mg_ce_id='"&mg_ce_id&"', mod_id='900001' where sido = '"&rs("sido")&"' and gugun = '"&rs("gugun")&"' and mg_group = '"&rs("mg_group")&"' "
	response.write(sql)
	dbconn.execute(sql)	  

	rs.movenext()
loop

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
end if

response.write("ó���Ǽ� : " + cstr(i))

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
%>