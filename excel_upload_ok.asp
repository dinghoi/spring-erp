<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon_db.asp" -->
<% 

	objFile = SERVER.MapPath(".") & "\srv_upload\주소록.xls"
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect
	
	sql_excel = "delete from up_excel"
	dbconn.execute(sql_excel)

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"
	
	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
	  if rowcount > -1 then
		for i=0 to rowcount
' 구군
		sql_etc = "select * from ce_area where sido = '" + xgr(6,i) +"' and gugun = '" + xgr(7,i) + "'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_gugun = tot_gugun + 1
			tot_err = tot_err + 1
			mg_ce_id = ""
		  else
			mg_ce_id = rs_etc("mg_ce_id")	  
		end if
' CE
		sql_etc = "select * from memb where user_id = '" + mg_ce_id + "'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_ce = tot_ce + 1
			tot_err = tot_err + 1
			mg_ce = "미등록"
		  else
			mg_ce = rs_etc("user_name")
		end if

		sql_excel="insert into up_excel (acpt_user,tel_ddd,tel_no1,tel_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce) values ('"&xgr(2,i)&"','"&xgr(3,i)&"','"&xgr(4,i)&"','"&xgr(5,i)&"','"&xgr(0,i)&"','"&xgr(1,i)&"','"&xgr(6,i)&"','"&xgr(7,i)&"','"&xgr(8,i)&"','"&xgr(9,i)&"','"&mg_ce_id&"','"&mg_ce&"')"
		dbconn.execute(sql_excel)

		next
	  end if
	err_msg = cstr(rowcount+1) + " 건 처리되었습니다..."
	response.write"<script language=javascript>"
	response.write"alert('"&err_msg&"');"
	response.write"history.go(-1);"
	response.write"</script>"
	Response.End

	rs.close
	cn.close
	rs_etc.close
	set rs = nothing
	set cn = nothing
	set rs_etc = nothing
%>