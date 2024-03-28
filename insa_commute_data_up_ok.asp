<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<% 
	on Error resume next

	objFile = request.form("objFile")

'	objFile = SERVER.MapPath(".") & "\srv_upload\주소록.xls"
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect
	
	dbconn.BeginTrans
	
  'Response.write objFile

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"
	
	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
    if rowcount > -1 then
		for i=0 to rowcount
		'Response.write xgr(0,i) 
		'Response.write xgr(1,i) 
		'Response.write mid(xgr(2,i), 1, 5) 
		'Response.write xgr(2,i)
		sql="insert into commute (emp_no,wrkt_dt,wrk_start_time,wrk_end_time,wrk_type) values "& _
		"('"&xgr(0,i)&"','"&xgr(1,i)&"','"&mid(xgr(2,i), 1, 5)&"','"&right(xgr(2,i), 5)&"','"&xgr(2,i)&"') on duplicate key update wrkt_dt='"&xgr(1,i)&"',wrk_start_time='"&mid(xgr(2,i), 1, 5)&"',wrk_end_time='"&right(xgr(2,i), 5)&"',wrk_type='"&xgr(2,i)&"'" ' wrk_type ='"&xgr(2,i)&"'"
		
		'Response.write sql
		dbconn.execute(sql)
		next
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		response.write"alert('"&Err.Description&"');"
		Response.End
	else    
		dbconn.CommitTrans
		end_msg = cstr(rowcount) +" 건 등록 완료되었습니다...."
	end if

	err_msg = cstr(rowcount+1) + " 건 처리되었습니다..."
	response.write"<script language=javascript>"
	response.write"alert('"&err_msg&"');"
	'response.write"location.replace('insa_commute_data_up.asp');"
	response.write"location.replace('insa_commute_mg.asp');"
	response.write"</script>"
	Response.End

	rs.close
	cn.close
	rs_etc.close
	set rs = nothing
	set cn = nothing
	set rs_etc = nothing
%>