<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	curr_date = mid(cstr(now()),1,10)
	car_old_no = request.form("car_old_no")
	
	del_date = cstr(mid(curr_date,1,4)) + cstr(mid(curr_date,6,2)) + cstr(mid(curr_date,9,2))

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
    Set Rs_org = Server.CreateObject("ADODB.Recordset")
    Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_car = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans


sql = "insert into car_info_del select '"&del_date&"' as car_del_date,car_info.* from car_info where car_no ='"&car_old_no&"'"
    dbconn.execute(sql)

sql = "delete from car_info where car_no ='"&car_old_no&"'"
    dbconn.execute(sql)	


	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
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
