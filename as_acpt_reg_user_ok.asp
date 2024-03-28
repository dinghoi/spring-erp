<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include file="xmlrpc.asp"-->
<!--#include file="class.EmmaSMS.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

' 최근수정 2010-06-28
	acpt_date = request.form("curr_date_time")
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
	dept = request.form("dept")
	sido = request.form("sido")
	gugun = request.form("gugun")
	dong = request.form("dong")
	addr = request.form("addr")
	mg_ce_id = request.form("mg_ce_id")
	mg_ce = request.form("mg_ce")
	team = request.form("team")
	reside_place = request.form("reside_place")	
	acpt_grade = user_grade

	if reside_place = "본사" then
		reside = "0"
	  else
	  	reside = "1"
	end if
	
	as_major = "PC/노트북"
	as_memo = request.form("as_memo")		
	as_device = request.form("as_device")
	maker = request.form("maker")
	model_no = request.form("model_no")
	request_date = request.form("request_date")
	request_hh = request.form("request_hh")
	request_mm = request.form("request_mm")	
	request_time = cstr(request_hh) + cstr(request_mm)
	as_process = "접수"
	as_type = request.form("as_type")
	regi_id = user_id
	view_ok = request.form("view_ok")
	serial_no = request.form("serial_no")
	asets_no = request.form("asets_no")
	w_cnt = 1
	acpt_hh = cstr(datepart("h", timevalue(acpt_date)))
	acpt_mm = cstr(datepart("n", timevalue(acpt_date)))
	acpt_ss = cstr(datepart("s", timevalue(acpt_date)))

	acpt_date = mid(acpt_date,1,10) + " " + acpt_hh + ":" + acpt_mm + ":" + acpt_ss 	

	'//2017-09-15 확인서 필수 여부 추가
	doc_yn = toString(request.form("doc_yn"),"N")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if user_id = "daekyo" then
		sms_yn = "N"
		mg_ce_id = "kwonhelp"
		mg_ce = "헬프"
	  else
		sms_yn = "Y"
	end if
	sms_msg = ""

	sql="select * from nkp.memb where user_id = '" + mg_ce_id + "'"
	set rs=dbconn.execute(sql)
	if	rs.eof or rs.bof then
		hp_no = "010-8737-2299"
		team = "error"
		reside_place = "error"
		reside_company = "error"
		reside = "E"
  	else
		hp_no = rs("hp")
		team = rs("team")
		reside_place = rs("reside_place")
		reside_company = rs("reside_company")
		reside = rs("reside")
	end if

	if	sms_yn = "Y" then
'		sms_msg = company + "-" + dept + "-" + acpt_user  + "-" + gugun + " " + dong + " " + addr
		sms_msg = company + "-" + dept + "-" + acpt_user  + "-" + gugun + " " + dong + " " + tel_ddd + tel_no1 + tel_no2
		sms_to = hp_no
'		sms_from = tel_ddd + tel_no1 + tel_no2
		sms_from = "1566-4711"
		sms_date = ""

		Set sms = new EmmaSMS
		sms.login "kwon5250", "kwon5230"	' sms.login [고객 ID], [고객 패스워드]
		ret = sms.send(sms_to, sms_from, sms_msg, sms_date)

		sms_msg = s_ce_id + "(" + s_ce + ") 님 핸드폰 " + hp_no + "으로 문자 발송 " 
		if ret = true then
			sms_yn = "Y"
		else
			sms_yn = "E"
			sms_msg = "문자전송 에라 !!!! "
		end if
	
		Set sms = Nothing
	end if		

	if s_team = "외주관리" then
		s_reside_place = rs("reside_place")
		mg_ce_id = s_ce_id
	end if	

'	if mg_group <> "1" then
'		w_cnt = 1
'	end if

	i = 0

	do until i = w_cnt
		i = i + 1
		
		sql="insert into nkp.as_acpt (acpt_date,acpt_man,acpt_grade,acpt_user,user_grade,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,maker,as_device,model_no,serial_no,asets_no,reside,reside_place,reside_company,team,sms,reg_id,doc_yn) values (cast('"&acpt_date&"' as datetime),'"&acpt_man&"','"&acpt_grade&"','"&acpt_user&"','"&user_grade&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&hp_ddd&"','"&hp_no1&"','"&hp_no2&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"','"&mg_ce&"','"&mg_group&"','"&as_memo&"','"&request_date&"','"&request_time&"','"&as_process&"','"&as_type&"','"&maker&"','"&as_device&"','"&model_no&"','"&serial_no&"','"&asets_no&"','"&reside&"','"&reside_place&"','"&reside_company&"','"&team&"','"&sms_yn&"','"&regi_id&"'"&",'"&doc_yn&"')"
		dbconn.execute(sql)
	loop
	                                       		
	sql="select * from nkp.juso_list where company='" + company + "' and dept = '" + dept + "'"
	set rs=dbconn.execute(sql)
	
	if	rs.eof or rs.bof then
		sql = "insert into nkp.juso_list (tel_ddd,tel_no1,tel_no2,company,dept,mg_group,sido,gugun,dong,addr,mg_ce_id,regi_date,regi_id,reside) values ('"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&company&"','"&dept&"','"&mg_group&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"',now(),'"&regi_id&"','"&reside&"')"
		dbconn.execute(sql)
	  else
		sql = "update nkp.juso_list set mg_group='"&mg_group&"', sido='"&sido&"', gugun='"&gugun&"', dong='"&dong&"', addr='"&addr&"', mg_ce_id='"&mg_ce_id&"' where company='" + company + "' and dept = '" + dept + "'"
		dbconn.execute(sql)	  
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + " " + cstr(w_cnt) +" 건 등록 완료되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('as_list_ce_user.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
