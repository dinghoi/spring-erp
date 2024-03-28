<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<% 

	dim acpt_no(10)
	dim visit_date(10)
	acpt_no(1) = request.form("acpt_no1")
	acpt_no(2) = request.form("acpt_no2")
	acpt_no(3) = request.form("acpt_no3")
	acpt_no(4) = request.form("acpt_no4")
	acpt_no(5) = request.form("acpt_no5")
	acpt_no(6) = request.form("acpt_no6")
	acpt_no(7) = request.form("acpt_no7")
	acpt_no(8) = request.form("acpt_no8")
	acpt_no(9) = request.form("acpt_no9")
	acpt_no(10) = request.form("acpt_no10")

	visit_date(1) = request.form("visit_date1")
	visit_date(2) = request.form("visit_date2")
	visit_date(3) = request.form("visit_date3")
	visit_date(4) = request.form("visit_date4")
	visit_date(5) = request.form("visit_date5")
	visit_date(6) = request.form("visit_date6")
	visit_date(7) = request.form("visit_date7")
	visit_date(8) = request.form("visit_date8")
	visit_date(9) = request.form("visit_date9")
	visit_date(10) = request.form("visit_date10")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	w_cnt = 0
	for i = 1 to 10
		
		if visit_date(i) <> "" then
			as_process = "완료"
			visit_time = "1300"		
			sql = "Update large_acpt set mg_ce_id='"+user_id+"',mg_ce='"+user_name+"',visit_date='"+visit_date(i)+"',visit_time='"+visit_time+ "',as_process='"+as_process
			sql = sql+"',as_memo='"+as_process+"',reside_place='"+reside_place+"',belong='"+belong+"',mod_date=now(),mod_id='"+mod_id+"',reside='"+reside+"' where acpt_no = "&int(acpt_no(i))		
			dbconn.execute(sql)
			w_cnt = w_cnt + 1
		end if
	next
	                                       		
	end_msg = sms_msg + " " + cstr(w_cnt) +" 건 등록 완료되었습니다...."
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('large_list_ce.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>