<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	on Error resume next

	acpt_man = request.form("acpt_man")
	acpt_user = request.form("acpt_user")
	user_grade = request.form("user_grade")
	tel_ddd = request.form("tel_ddd")
	tel_no1 = request.form("tel_no1")
	tel_no2 = request.form("tel_no2")
	hp_ddd = request.form("hp_ddd")
	hp_no1 = request.form("hp_no1")
	hp_no2 = request.form("hp_no2")
	if hp_no1 = "" then
		hp_ddd = ""
		hp_no1 = ""
		hp_no2 = ""
	end if
	company = request.form("company")
	org_first = request.form("org_first")
	org_second = request.form("org_second")
	dept_name = request.form("dept_name")
	dept = org_first + " " + dept_name
	if isnull(dept_name) then
		dept = org_first + " " + org_second
	end if
	dept_code = request.form("dept_code")
	internet_no = request.form("internet_no")
	as_sw = request.form("as_sw")
	if as_sw = "Y" then
		sido = request.form("old_sido")
		gugun = request.form("old_gugun")
		dong = request.form("old_dong")
		addr = request.form("old_addr")
		mg_ce_id = request.form("old_mg_ce_id")
		mg_ce = request.form("old_mg_ce")
		team = request.form("old_team")
		reside_place = request.form("old_reside_place")
	  else	
		sido = request.form("sido")
		gugun = request.form("gugun")
		dong = request.form("dong")
		addr = request.form("addr")
		mg_ce_id = request.form("mg_ce_id")
		mg_ce = request.form("mg_ce")
		team = request.form("team")
		reside_place = request.form("reside_place")
	end if
	
	as_memo = request.form("as_memo")		
	as_device = request.form("as_device")
	maker = request.form("maker")
	serial_no = request.form("serial_no")
	model_no = "."
	request_date = request.form("request_date")
	request_hh = request.form("request_hh")
	request_mm = request.form("request_mm")	
	request_time = cstr(request_hh) + cstr(request_mm)
	as_process = "접수"
	as_type = "원격처리"
	if as_sw = "N" then
		as_type = "이전설치"
		if internet_no <> "" then
			as_memo = "인터넷 이전 신청 필요 인터넷번호 ( " + internet_no + " ) " + as_memo
		end if
	end if
	sms_yn = "N"
	
	if asset_company = "00" then
		asset_company = "01"
	end if

	curr_date = mid(cstr(now()),1,10)
	curr_hh = int(cstr(datepart("h",now)))
	curr_mm = int(cstr(datepart("n",now)))
	request_date = curr_date
	request_hh = curr_hh
	request_mm = curr_mm
	
	if curr_hh < 10 then
		curr_hh = "0" + cstr(curr_hh)
	end if
	
	if curr_mm < 10 then
		curr_mm = "0" + cstr(curr_mm)
	end if
	
	if request_mm < "30" then
		request_mm = "30"
	end if
	
	if request_mm > "30" then
		request_mm = "00"
		request_hh = cstr(request_hh + 1)
	end if
	
	request_hh = cstr(request_hh + 4)
	
	if request_hh = "18" then
		request_mm = "00"
	end if
	
	if request_hh > "18" then
		request_hh = request_hh - 18
		request_date = mid(cstr(now()+1),1,10)
		select case request_hh
			case 1
				request_hh = "10"
			case 2
				request_hh = "11"
			case 3
				request_hh = "12"
			case else
				request_hh = "13"
		end select	
	end if
	
	c_w = datepart("w",curr_date)
	
	if c_w = 7 or c_w = 1 then
		request_hh = "13"
		request_mm = "00"
	end if

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	
	dbconn.BeginTrans

	for k = 1 to 15
	
		w = datepart("w",request_date)
	
		if w = 7 then
			request_date = dateadd("d",2,request_date)
		end if
		
		if w = 1 then
			request_date = dateadd("d",1,request_date)
		end if
	'response.write(w)
		Set Rs_hol = Server.CreateObject("ADODB.Recordset")
		Sql="select * from holiday where holiday = '"&request_date&"'"
		Rs_hol.Open Sql, Dbconn, 1
		if 	rs_hol.eof then
			request_date = request_date
			exit for
		else
			request_date = dateadd("d",1,request_date)
		end if
	
		k = k + 1
	'	rs_hol.Close()
	next
	rs_hol.Close()
	request_time = cstr(request_hh) + cstr(request_mm)

	sql="select * from memb where user_id = '" + mg_ce_id + "'"
	set rs=dbconn.execute(sql)
	if	rs.eof or rs.bof then
		reside = "0"
  	else
		reside = rs("reside")
	end if

	sql="insert into as_acpt (acpt_date,acpt_man,acpt_grade,acpt_user,user_grade,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,maker,as_device,serial_no,reside,reside_place,team,sms,reg_id) values (now(),'"&acpt_man&"','"&acpt_grade&"','"&acpt_user&"','"&user_grade&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&hp_ddd&"','"&hp_no1&"','"&hp_no2&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"','"&mg_ce&"','"&mg_group&"','"&as_memo&"','"&request_date&"','"&request_time&"','"&as_process&"','"&as_type&"','"&maker&"','"&as_device&"','"&serial_no&"','"&reside&"','"&reside_place&"','"&team&"','"&sms_yn&"','"&user_id&"')"
	dbconn.execute(sql)
	                                       		
	sql = "update asset_dept set sido='"&sido&"', gugun='"&gugun&"', dong='"&dong&"', addr='"&addr&"', person='"&acpt_user&"', tel_ddd='"&tel_ddd&"', tel_no1='"&tel_no1&"', tel_no2='"&tel_no2&"' where company='" + asset_company + "' and dept_code = '" + dept_code + "'"
	dbconn.execute(sql)	  

	sql="select * from juso_list where company='" + company + "' and dept = '" + dept + "'"
	set rs=dbconn.execute(sql)
	
	if	rs.eof or rs.bof then
		sql = "insert into juso_list (tel_ddd,tel_no1,tel_no2,company,dept,mg_group,sido,gugun,dong,addr,mg_ce_id,regi_date,regi_id,reside) values ('"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&company&"','"&dept&"','"&mg_group&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"',now(),'"&regi_id&"','"&reside&"')"
		dbconn.execute(sql)
	  else
		sql = "update juso_list set mg_group='"&mg_group&"', sido='"&sido&"', gugun='"&gugun&"', dong='"&dong&"', addr='"&addr&"', mg_ce_id='"&mg_ce_id&"' where company='" + company + "' and dept = '" + dept + "'"
		dbconn.execute(sql)	  
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록 완료되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('nh_as_reg.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
