<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
' 대량 데이터 batch upload

Dim DbConnect
DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=Wlsgustn6!;"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

Dbconn.BeginTrans

sql = "select * from nkp.large_acpt" 
Rs.Open Sql, Dbconn, 1 

paper_no = "2014081101"

do until rs.eof

	sql = "insert into as_acpt(acpt_date,acpt_man,acpt_grade,acpt_user,user_grade,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept"& _					
	",sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,maker,as_device,model_no,serial_no"& _
	",asets_no,reside_place,team,large_paper_no) values "& _
 	" (now(),'"&rs("acpt_man")&"','"&rs("acpt_grade")&"','"&rs("acpt_user")&"','"&rs("user_grade")&"','"&rs("tel_ddd")&"','"&rs("tel_no1")& _
	"','"&rs("tel_no2")&"','"&rs("hp_ddd")&"','"&rs("hp_no1")&"','"&rs("hp_no2")&"','"&rs("company")&"','"&rs("dept")&"','"&rs("sido")& _
	"','"&rs("gugun")&"','"&rs("dong")&"','"&rs("addr")&"','"&rs("mg_ce_id")&"','"&rs("mg_ce")&"','"&rs("mg_group")&"','"&rs("as_memo")& _
	"','"&rs("request_date")&"','"&rs("request_time")&"','"&rs("as_process")&"','"&rs("as_type")&"','"&rs("maker")&"','"&rs("as_device")& _
	"','"&rs("model_no")&"','"&rs("serial_no")&"','"&rs("asets_no")&"','"&rs("reside_place")&"','"&rs("belong")& _
	"','"&paper_no&"')"	

	dbconn.execute(sql)

	rs.movenext
loop

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
end if

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
%>