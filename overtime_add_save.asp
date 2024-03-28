<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	work_item = request.form("work_item")
	work_date = request.form("work_date")
	old_date = request.form("old_date")
	company = request.form("company")
	dept = request.form("dept")
	from_hh = request.form("from_hh")
	from_mm = request.form("from_mm")	
	from_time = cstr(from_hh) + cstr(from_mm)
	to_hh = request.form("to_hh")
	to_mm = request.form("to_mm")	
	to_time = cstr(to_hh) + cstr(to_mm)
	work_gubun_amt = request.form("work_gubun")
	i=instr(1,work_gubun_amt,"-")'
	work_gubun = mid(work_gubun_amt,1,i-1)
	overtime_amt = int(mid(work_gubun_amt,i+1))
	mg_ce_id = request.form("mg_ce_id")
	work_memo = request.form("work_memo")
	sign_no = request.form("sign_no")
	you_yn = request.form("you_yn")
	cancel_yn = request.form("cancel_yn")
	cost_detail = work_gubun

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "select * from etc_code where etc_type = '41' and etc_name = '"&work_gubun&"'"
	set rs_etc=dbconn.execute(sql)
	if rs_etc.eof or rs_etc.bof then
		cost_detail = "ERROR"
	  else		
		cost_detail = rs_etc("group_name")
	end if

	sql = "select * from memb where user_id = '"&user_id&"'"
	set rs_memb=dbconn.execute(sql)

'	if	u_type = "U" then
		sql = "Update overtime set from_time='"&from_time&"',to_time='"&to_time&"',work_gubun='"&work_gubun&"',cost_detail='"&cost_detail& _
		"',overtime_amt='"&overtime_amt&"',work_memo='"&work_memo&"',sign_no='"&sign_no&"',you_yn='"&you_yn&"',cancel_yn='"&cancel_yn& _
		"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where work_date = '"&work_date&"' and mg_ce_id = '"&mg_ce_id&"'"
		dbconn.execute(sql)
'	  else
'		sql="insert into overtime (work_date,mg_ce_id,acpt_no,emp_company,bonbu,saupbu,team,org_name,company,dept,work_item,from_time,to_time"& _
'		",work_gubun,overtime_amt,work_memo,cancel_yn,end_sw,reg_id,reg_date) values ('"&work_date&"','"&mg_ce_id&"',"&acpt_no& _
'		",'"&rs_memb("emp_company")&"','"&rs_memb("bonbu")&"','"&rs_memb("saupbu")&"','"&rs_memb("team")&"','"&rs_memb("org_name")& _
'		"','"&company&"','"&dept&"','"&work_item&"','"&from_time&"','"&to_time&"','"&work_gubun&"',"&overtime_amt&",'"&work_memo& _
'		"','"&cancel_yn&"','"&end_sw&"','"&user_id&"',now())"
'		dbconn.execute(sql)
'	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "처리 되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
