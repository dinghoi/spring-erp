<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

    emp_user = request.cookies("nkpmg_user")("coo_user_name")
    emp_no = request.cookies("nkpmg_user")("coo_user_id")

    pe_car_no=Request.form("pe_car_no")
    pe_date=Request.form("pe_date")
    pe_seq=Request.form("pe_seq")
    car_name=Request.form("car_name")

	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_pe = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_trans = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans

    sql = " delete from car_penalty " & _
	            "  where pe_car_no = '"&pe_car_no&"' and pe_date = '"&pe_date&"' and pe_seq = '"&pe_seq&"'"
	
	dbconn.execute(sql)
	
	
    if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if

'	if view_condi = "성명" then 
'	           url = "insa_master_modify.asp?ck_sw=y&condi=" + emp_name + "&view_condi="+ view_condi
'		else 
'		       url = "insa_master_modify.asp?ck_sw=y&condi=" + emp_no + "&view_condi="+ view_condi
'	end if
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('insa_car_penalty_list.asp');"
'	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
