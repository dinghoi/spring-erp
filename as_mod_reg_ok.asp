<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include file="xmlrpc.asp"-->
<!--#include file="class.EmmaSMS.asp"-->
<%
'	on Error resume next

	page = request.form("page")
	from_date = request.form("from_date")
	to_date = request.form("to_date")
	date_sw = request.form("date_sw")
	process_sw = request.form("process_sw")
	field_check = request.form("field_check")
	field_view = request.form("field_view")
	condi_com = request.form("condi_com")

	acpt_no = request.form("acpt_no")
	acpt_user = request.form("acpt_user")
	user_grade = request.form("user_grade")
	tel_ddd = request.form("tel_ddd")
	tel_no1 = request.form("tel_no1")
	tel_no2 = request.form("tel_no2")
	hp_ddd = request.form("hp_ddd")
	hp_no1 = request.form("hp_no1")
	hp_no2 = request.form("hp_no2")
	company = request.form("company")
	dept = request.form("dept")		
	sido = request.form("sido")
	gugun = request.form("gugun")
	dong = request.form("dong")
	addr = request.form("addr")
	mg_ce_id = request.form("mg_ce_id")
	mg_ce = request.form("mg_ce")
	reside_place = request.form("reside_place")
	ce_mod_ck = request.form("ce_mod_ck")
	if	ce_mod_ck = "1" then		
		mg_ce_id = request.form("s_ce_id")
		mg_ce = request.form("s_ce")
		reside_place = request.form("s_reside_place")
	  else
	  	ce_mod_ck = "0"
	end if
	sms_yn = request.form("sms_yn")
	sms_old = request.form("sms_old")
	as_memo = request.form("as_memo")
	request_date = request.form("request_date")
	request_hh = request.form("request_hh")
	request_mm = request.form("request_mm")
	as_type = request.form("as_type")
	request_time = cstr(request_hh) + cstr(request_mm)
	sms_yn = request.form("sms_yn")
	
	cowork_yn = request.form("cowork_yn")
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sms_msg = ""
	sql="select * from memb where user_id = '" + mg_ce_id + "'"
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

		sms_msg = mg_ce_id + "(" + mg_ce + ") 님 핸드폰 " + hp_no + "으로 문자 발송 " 
		if ret = true then
			sms_yn = "Y"
		else
			sms_msg = "문자전송 에라 !!!! "
			sms_yn = "E"
		end if
	
		Set sms = Nothing
	end if		
	if sms_old <> "Y" then
		sms_old = sms_yn
	end if

	sql = "UPDATE as_acpt "&_
	      "   SET acpt_user       = '"&acpt_user&"'      "&_
	      "      , user_grade     = '"&user_grade&"'     "&_
	      "      , tel_ddd        = '"&tel_ddd&"'        "&_
	      "      , tel_no1        = '"&tel_no1&"'        "&_
	      "      , tel_no2        = '"&tel_no2&"'        "&_
	      "      , hp_ddd         = '"&hp_ddd&"'         "&_
	      "      , hp_no1         = '"&hp_no1&"'         "&_
	      "      , hp_no2         = '"&hp_no2&"'         "&_
	      "      , company        = '"&company&"'        "&_
	      "      , dept           = '"&dept&"'           "&_
	      "      , sido           = '"&sido&"'           "&_
	      "      , gugun          = '"&gugun&"'          "&_
	      "      , dong           = '"&dong&"'           "&_
	      "      , addr           = '"&addr&"'           "&_
	      "      , mg_ce_id       = '"&mg_ce_id&"'       "&_
	      "      , mg_ce          = '"&mg_ce&"'          "&_
	      "      , request_date   = '"&request_date&"'   "&_
	      "      , request_time   = '"&request_time&"'   "&_
	      "      , as_memo        = '"&as_memo&"'        "&_
	      "      , as_type        = '"&as_type&"'        "&_
	      "      , team           = '"&team&"'           "&_
	      "      , reside         = '"&reside&"'         "&_
	      "      , reside_company = '"&reside_company&"' "&_
	      "      , reside_place   = '"&reside_place&"'   "&_
	      "      , mod_date       = now()                "&_
	      "      , mod_id         = '"&user_id&"'        "&_
	      "      , sms            = '"&sms_old&"'        "&_
	      "      , cowork_yn      = '"&cowork_yn&"'      "&_
	      " WHERE  acpt_no        = "&int(acpt_no)
	dbconn.execute(sql)

	sql="SELECT * FROM juso_list WHERE company='" + company + "' AND dept = '" + dept + "'"
	set rs=dbconn.execute(sql)
	
	if	rs.eof or rs.bof then
		sql = "INSERT INTO juso_list (tel_ddd,tel_no1,tel_no2,company,dept,mg_group,sido,gugun,dong,addr,mg_ce_id,regi_date,regi_id) "&_
		      " VALUES ('"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&company&"','"&dept&"','"&mg_group&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"',now(),'"&user_id&"')"
		dbconn.execute(sql)
	  else
		sql = "UPDATE juso_list SET tel_ddd='"+tel_ddd+"',tel_no1='"+tel_no1+"',tel_no2='"+tel_no2+"',sido='"+sido+"',gugun='"+gugun+"',dong='"+dong+"',addr='"+addr+"' where company='" + company + "' and dept = '" + dept + "'"
		dbconn.execute(sql)	  
	end if
' 변경 History 저장
	mod_pg = "데이터수정"
	sql = "INSERT INTO as_mod (acpt_no, mod_date, mod_id, mod_name, mod_pg) values ("&int(acpt_no)&", now(), '"&user_id&"', '"&user_name&"', '"&mod_pg&"')"
	dbconn.execute(sql)

	check_sw = "y"
	url = "as_list.asp?page="+page+"&from_date="+from_date+"&to_date="+to_date+"&date_sw="+date_sw+"&process_sw="+process_sw+"&field_check="+field_check+"&field_view="+field_view+"&ck_sw="+check_sw+"&company="+condi_com

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "변경되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"location.replace('"&url&"');"
'	response.write"alert('변경되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
