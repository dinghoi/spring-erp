<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	
'	on Error resume next

	asset_no = request.form("asset_no")
	serial_no = request.form("serial_no")
	company_name = request.form("company_name")
	gubun = request.form("gubun")
	maker = request.form("maker")
	dept_code = request.form("dept_code")
	old_code = request.form("old_code")
	dept_name = request.form("dept_name")
	old_name = request.form("old_name")
	user_name = request.form("user_name")
	old_user = request.form("old_user")
	install_date = request.form("install_date")
	request_hh = request.form("request_hh")
	request_mm = request.form("request_mm")	
	request_time = cstr(request_hh) + cstr(request_mm)
	old_date = request.form("old_date")
	trans_memo = request.form("trans_memo")
	as_memo = "전화 연락후 시간 약속을 한 후 이전설치 해주시길 바랍니다."
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	
	dbconn.BeginTrans

	etc_code = "79" + cstr(gubun)
	sql = "select * from etc_code where etc_code ='" + etc_code + "'"
	set rs_etc=dbconn.execute(sql)
	as_device = rs_etc("etc_name")

	sql = "select * from asset_dept where company ='" + mid(asset_no,1,2) + "' and dept_code ='" + dept_code + "'"
	set rs=dbconn.execute(sql)
	org_first = rs("org_first")
	org_second = rs("org_second")
	dept_name = rs("dept_name")
	dept = org_first + " " + dept_name
	if isnull(dept_name) then
		dept = org_first + " " + org_second
	end if

	sql = "select * from ce_area where sido = '" + rs("sido") + "' and gugun = '" + rs("gugun") + "' and mg_group = '" + mg_group + "'"
	set rs_ce=dbconn.execute(sql)
	mg_ce_id = rs_ce("mg_ce_id")
	sql = "select * from memb where user_id = '" + mg_ce_id + "'"
	set rs_memb=dbconn.execute(sql)
	mg_ce = rs_memb("user_name")
	team = rs_memb("team")
	reside = rs_memb("reside")
	reside_place = rs_memb("reside_place")

	sql="insert into as_acpt (acpt_date,acpt_man,acpt_grade,acpt_user,tel_ddd,tel_no1,tel_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,maker,as_device,serial_no,asets_no,reside,reside_place,team,sms,mod_id,mod_date) values (now(),'인터넷','회사','"&rs("person")&"','"&rs("tel_ddd")&"','"&rs("tel_no1")&"','"&rs("tel_no2")&"','"&company_name&"','"&dept&"','"&rs("sido")&"','"&rs("gugun")&"','"&rs("dong")&"','"&rs("addr")&"','"&mg_ce_id&"','"&mg_ce&"','"&mg_group&"','"&as_memo&"','"&install_date&"','"&request_time&"','접수','이전설치','"&maker&"','"&as_device&"','"&serial_no&"','"&asset_no&"','"&reside&"','"&reside_place&"','"&team&"','N','"&user_id&"',now())"
	dbconn.execute(sql)

	sql = "Update asset set dept_code ='"+dept_code+"', user_name ='"+user_name+"', install_date='"+install_date+"', mod_id='"+user_id+"', mod_date=now() where asset_no = '" + asset_no + "'"
	dbconn.execute(sql)

	sql = "select asset_no, max(history_seq) as max_seq from asset_history where asset_no='" + asset_no + "'"
	set rs_hist=dbconn.execute(sql)
	
	if	isnull(rs_hist("max_seq"))  then
		history_seq = 1
	  else
		history_seq = cint(rs_hist("max_seq")) + 1
	end if

	sql="insert into asset_history (asset_no,history_seq,dept_code,user_name,install_date,trans_memo,reg_id,reg_date) values ('"&asset_no&"',"&history_seq&",'"&old_code&"','"&old_user&"','"&old_date&"','"&trans_memo&"','"&user_id&"',now())"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "이전 설치 저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "이전 설치 완료되었습니다...."
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

