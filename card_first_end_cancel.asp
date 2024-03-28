<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	slip_month=Request.form("slip_month")
	card_type=Request.form("card_type")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")

	from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

'마감
'	sql = "select * from card_slip where slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"'"
'	response.write(sql)
'	Rs.Open Sql, Dbconn, 1

'	do until rs.eof
		sql = "Update card_slip set account_end='N' where slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"'"
		dbconn.execute(sql)
'		rs.movenext()
'	loop
'	rs.close()
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "마감처리 되었습니다...."
	end if

	url = "card_slip_mg.asp?slip_month="&slip_month&"&card_type="&card_type&"&field_check="&field_check&"&field_view="&field_view&"&ck_sw="&"y"
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


