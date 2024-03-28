<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkp_itft_db.asp" -->
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
	acpt_grade = user_grade
'	mg_group = request.form("mg_group")
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
'	team = request.form("team")
'	reside_place = request.form("reside_place")	
'	reside_company = request.form("reside_company")	
	s_ce_id = request.form("s_ce_id")
	s_ce = request.form("s_ce")
'	s_team = request.form("s_team")
'	s_reside_place = request.form("s_reside_place")	
'	s_reside_company = request.form("s_reside_company")	
'	s_reside_place = request.cookies("nkpmg_user")("coo_reside_place")
'	s_reside = cstr(request.cookies("nkpmg_user")("coo_reside"))

	if s_ce_id = "" or s_ce_id < "0" then
		s_ce_id = mg_ce_id
		s_ce = mg_ce
'		s_team = team
'		s_reside_place = reside_place
'		s_reside_company = reside_company
	end if

'	if s_reside_place = "본사" then
'		reside = "0"
'	  else
'	  	reside = "1"
'	end if
	
	as_major = "PC/노트북"
	as_memo = request.form("as_memo")		
	as_memo = Replace(as_memo,"'","&quot;")
	as_device = request.form("as_device")
	maker = request.form("maker")
	model_no = request.form("model_no")
	request_date = request.form("request_date")
	request_hh = request.form("request_hh")
	request_mm = request.form("request_mm")	
	request_time = cstr(request_hh) + cstr(request_mm)
	as_process = "접수"
	as_type = request.form("as_type")
	visit_request_yn = request.form("visit_request")
	if as_type <> "방문처리" or visit_request_yn = "" or isnull(visit_request_yn) then
		visit_request_yn = "N"
	end if

	regi_id = user_id
	view_ok = request.form("view_ok")
	serial_no = request.form("serial_no")
	asets_no = request.form("asets_no")
	w_cnt = int(request.form("w_cnt"))
	acpt_hh = cstr(datepart("h", timevalue(acpt_date)))
	acpt_mm = cstr(datepart("n", timevalue(acpt_date)))
	acpt_ss = cstr(datepart("s", timevalue(acpt_date)))

	acpt_date = mid(acpt_date,1,10) + " " + acpt_hh + ":" + acpt_mm + ":" + acpt_ss 	

	dbconn.BeginTrans

	sms_yn = request.form("sms_yn")
	sms_msg = ""

	'//2017-09-15 확인서 필수 여부 추가
	doc_yn = Trim(request.form("doc_yn")&"")
	If doc_yn<>"Y" Then doc_yn = "N" End IF

	sql="select * from nkp.memb where user_id = '" + s_ce_id + "'"
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

'	if s_team = "외주관리" then
'		s_reside_place = rs("reside_place")
'		mg_ce_id = s_ce_id
'	end if	

'	if mg_group <> "1" then
'		w_cnt = 1
'	end if

	i = 0

	do until i = w_cnt
		i = i + 1
		
		if	(company = "케") then		
			sql="insert into nkp.as_acpt (acpt_date,acpt_man,acpt_grade,acpt_user,user_grade,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,visit_request_yn,maker,as_device,model_no,serial_no,asets_no,reside,reside_place,reside_company,team,sms,reg_id,write_date,write_cnt,doc_yn) values (cast('"&acpt_date&"' as datetime),'"&acpt_man&"','"&acpt_grade&"','"&acpt_user&"','"&user_grade&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&hp_ddd&"','"&hp_no1&"','"&hp_no2&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&s_ce_id&"','"&s_ce&"','"&mg_group&"','"&as_memo&"','"&request_date&"','"&request_time&"','"&as_process&"','"&as_type&"','"&visit_request_yn&"','"&maker&"','"&as_device&"','"&model_no&"','"&serial_no&"','"&asets_no&"','"&reside&"','"&reside_place&"','"&reside_companye&"','"&team&"','"&sms_yn&"','"&regi_id&"',cast('"&acpt_date&"' as datetime),"&i&",'"&doc_yn&"')"
			dbconn.execute(sql)
			'Response.write sql
			itft_type = as_type
			if as_type = "방문처리" then
				itft_type = "A/S방문"
			end if
			if as_type = "신규설치" or as_type = "신규설치공사" or as_type = "이전설치" or as_type = "이전설치공사" or as_type = "장비회수" then
				itft_type = "설치"
			end if
			sql="insert into nkp.as_acpt (acpt_date,acpt_man,acpt_user,user_man,tel_ddd,tel_no1,tel_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,as_major,as_memo,request_date,request_time,as_process,as_type,maker,as_device,model_no,serial_no,asets_no,reside_place,sms,mg_company,write_date,write_cnt,doc_yn) values (cast('"&acpt_date&"' as datetime),'"&acpt_man&"','"&acpt_user&"','"&acpt_user&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&s_ce_id&"','"&s_ce&"','"&as_major&"','"&as_memo&"','"&request_date&"','"&request_time&"','"&as_process&"','"&itft_type&"','"&maker&"','"&as_device&"','"&model_no&"','"&serial_no&"','"&asets_no&"','콜센터','"&sms_yn&"','케이원',cast('"&acpt_date&"' as datetime),"&i&",'"&doc_yn&"')"
			'Response.write(sql)
			dbconn1.execute(sql)
		else
			sql="insert into nkp.as_acpt (acpt_date,acpt_man,acpt_grade,acpt_user,user_grade,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,visit_request_yn,maker,as_device,model_no,serial_no,asets_no,reside,reside_place,reside_company,team,sms,reg_id,doc_yn) values (cast('"&acpt_date&"' as datetime,doc_yn),'"&acpt_man&"','"&acpt_grade&"','"&acpt_user&"','"&user_grade&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&hp_ddd&"','"&hp_no1&"','"&hp_no2&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&s_ce_id&"','"&s_ce&"','"&mg_group&"','"&as_memo&"','"&request_date&"','"&request_time&"','"&as_process&"','"&as_type&"','"&visit_request_yn&"','"&maker&"','"&as_device&"','"&model_no&"','"&serial_no&"','"&asets_no&"','"&reside&"','"&reside_place&"','"&reside_company&"','"&team&"','"&sms_yn&"','"&user_id&"','"&doc_yn&"')"
			'Response.write sql
			dbconn.execute(sql)
		end if  
	loop
	                                       		
	sql="select * from nkp.juso_list where company='" + company + "' and dept = '" + dept + "'"
	set rs=dbconn.execute(sql)
	
	if	rs.eof or rs.bof then
		sql = "insert into nkp.juso_list (tel_ddd,tel_no1,tel_no2,company,dept,mg_group,sido,gugun,dong,addr,mg_ce_id,regi_date,regi_id,reside) values ('"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&company&"','"&dept&"','"&mg_group&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"',now(),'"&regi_id&"','"&reside&"')"
		'Response.write sql
		dbconn.execute(sql)
	  else
		sql = "update nkp.juso_list set mg_group='"&mg_group&"', sido='"&sido&"', gugun='"&gugun&"', dong='"&dong&"', addr='"&addr&"', mg_ce_id='"&mg_ce_id&"' where company='" + company + "' and dept = '" + dept + "'"
		'Response.write sql
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
	response.write"location.replace('as_list_ce.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
