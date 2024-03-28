<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/mysql_schema.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include file="xmlrpc.asp"-->
<!--#include file="class.EmmaSMS.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)

	page = abc("page")
	from_date = abc("from_date")
	to_date = abc("to_date")
	date_sw = abc("date_sw")
	process_sw = abc("process_sw")
	field_check = abc("field_check")
	field_view = abc("field_view")
	view_sort = abc("view_sort")
	page_cnt = abc("page_cnt")
	condi_com = abc("condi_com")
	view_c = abc("view_c")

	mg_group = abc("mg_group")
	acpt_no = abc("acpt_no")
	acpt_user = abc("acpt_user")
	tel_ddd = abc("tel_ddd")
	tel_no1 = abc("tel_no1")
	tel_no2 = abc("tel_no2")
	company = abc("company")
	dept = abc("dept")		
	sido = abc("sido")
	gugun = abc("gugun")
	dong = abc("dong")
	addr = abc("addr")
	mg_ce_id = abc("mg_ce_id")
	mg_ce = abc("mg_ce")
	ce_mod_ck = abc("ce_mod_ck")
	if	ce_mod_ck = "1" then		
		mg_ce_id = abc("s_ce_id")
		mg_ce = abc("s_ce")
		team = abc("s_team")
	'	reside_place = abc("s_reside_place")
	  else
	  	ce_mod_ck = "0"
	end if

	reside_place = abc("reside_place")
	if team = "외주관리" then
		reside_place = abc("s_reside_place")
	end if
			
'	sms_yn = abc("sms_yn")
	as_memo = abc("as_memo")
	request_date = abc("request_date")
	request_hh = abc("request_hh")
	request_mm = abc("request_mm")
	visit_date = abc("visit_date")
	visit_hh = abc("visit_hh")
	visit_mm = abc("visit_mm")
	as_process = abc("as_process")
	as_process_old = abc("as_process_old")
	as_type = abc("as_type")
	as_type_old = abc("as_type_old")
	into_reason = abc("into_reason")
	maker = abc("maker")
	as_device = abc("as_device")
	asets_no = abc("asets_no")
	serial_no = abc("serial_no")
	model_no = abc("model_no")
	as_parts = abc("as_parts")
	as_history = abc("as_history")
	if as_type = "방문처리" or as_type = "원격처리" or as_type = "기타" then
		dev_inst_cnt = 1
		ran_cnt = 0
		work_man_cnt = 1
		alba_cnt = 0
		person_amt = 1
	  else
		dev_inst_cnt = int(abc("dev_inst_cnt"))
		ran_cnt = int(abc("ran_cnt"))
		work_man_cnt = int(abc("work_man_cnt"))
		alba_cnt = int(abc("alba_cnt"))
		person_amt = dev_inst_cnt + ran_cnt
	end if
	write_date = abc("write_date")
	write_cnt = abc("write_cnt")
	request_time = cstr(request_hh) + cstr(request_mm)
	visit_time = cstr(visit_hh) + cstr(visit_mm)

' 입고 관리	
	acpt_no = int(acpt_no)
	in_seq = 1
	in_date = abc("in_date")
	in_process = "-"
	in_place = abc("in_place")
	in_replace = abc("in_replace")
	
	sms_yn = abc("new_sms")
	response.write(sms_yn)
	sms_old = abc("sms_old")
'	if sms_old = "Y" then
'		sms_yn = "Y"
'	end if
	be_pg = abc("be_pg")
	
	if as_process = "접수" then
'		arrival_date = a_null
'		asets_no = ""
'		serial_no = ""
		as_parts = ""
		as_history = ""
	end if
	
	err01 = ""
	err02 = ""
	err03 = ""
	err04 = ""
	err05 = ""
	err06 = ""
	err07 = ""
	err09 = ""

	if as_process = "완료" or as_process = "취소" or as_process = "대체" or as_process = "대체입고" then
		select case as_device
			case "데스크탑", "노트북", "DTO", "DTS"
				err01 = abc("err01")	
				err02 = abc("err02")					
			case "모니터"
				err03 = abc("err03")	
			case "프린터", "스케너", "플로터"
				err04 = abc("err04")	
			case "통신장비", "AP", "허브", "라우터", "TA", "네트웍장비", "회선"
				err05 = abc("err05")
			case "서버", "워크스테이션"
				err06 = abc("err06")
			case "아답터"
				err07 = abc("err07")
			case "기타"
				err09 = abc("err09")	
		end select												
	end if
'	reside_place = abc("reside_place")
'	team = abc("team")
'	if reside_place = "본사" then
'		reside = "0"
'	  else
'	  	reside = "1"
'	end if
	
	mod_id = user_id

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if as_process = "입고" or as_process = "대체입고" then
		sql = "select * from nkp.as_into where acpt_no="&acpt_no&" and in_seq="&in_seq
		set rs = dbconn.execute(sql)
		
		if	rs.eof or rs.bof then
			sql="insert into nkp.as_into (acpt_no,in_seq,in_process,into_date,in_place,in_remark,reg_id,reg_name,reg_date) values ('"&acpt_no&"','"&in_seq&"','"&in_process&"','"&in_date&"','"&in_place&"','"&into_reason&"','"&user_id&"','"&user_name&"',now())"
			dbconn.execute(sql)
		  else
			response.write"<script language=javascript>"
			response.write"alert('이미 입고처리가 되어 있습니다....');"		
	'		response.write"location.replace('as_list_ce.asp');"
			response.write"history.go(-1);"
			response.write"</script>"
			Response.End
		end if
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
		org_name = "error"
  	else
		hp_no = rs("hp")
		team = rs("team")
		reside_place = rs("reside_place")
		reside_company = rs("reside_company")
		reside = rs("reside")
		org_name = rs("org_name")
	end if

'	if sms_old <> "Y" then
		if	sms_yn = "Y" then
'			sms_msg = company + "-" + dept + "-" + acpt_user  + "-" + gugun + " " + dong + " " + addr
			sms_msg = company + "-" + dept + "-" + acpt_user  + "-" + gugun + " " + dong + " " + tel_ddd + tel_no1 + tel_no2
			sms_to = hp_no
'			sms_from = tel_ddd + tel_no1 + tel_no2
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
'	end if
'2014-01-25 기존에 설치사진 첨부 (시작)
'2014-07-12폴더 생성
	path_nm = "D:\web\att_file\" + company

    Set fso=Server.CreateObject("Scripting.FileSystemObject")'
	if Not fso.FolderExists(path_nm) then
		path_nm = fso.CreateFolder(path_nm)
	end if
	Set fso = Nothing

	path_name = "/att_file/" + company
	path = Server.MapPath (path_name)

	Set filenm1 = abc("att_file1")(1)
	filename1 = filenm1
	if filenm1 <> "" then 
		filename1 = filenm1.safeFileName	
		fileType1 = mid(filename1,inStrRev(filename1,".")+1)
		filename1 = company + "_" + sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_1." + fileType1
		save_path1 = path & "\" & filename1
	end if
	
	Set filenm2 = abc("att_file2")(1)
	filename2 = filenm2
	if filenm2 <> "" then 
		filename2 = filenm2.safeFileName	
		fileType2 = mid(filename2,inStrRev(filename2,".")+1)
		filename2 = company + "_" + sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_2." + fileType2
		save_path2 = path & "\" & filename2
	end if
	
	Set filenm3 = abc("att_file3")(1)
	filename3 = filenm3
	if filenm3 <> "" then 
		filename3 = filenm3.safeFileName	
		fileType3 = mid(filename3,inStrRev(filename3,".")+1)
		filename3 = company + "_" + sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_3." + fileType3
		save_path3 = path & "\" & filename3
	end if
	
	Set filenm4 = abc("att_file4")(1)
	filename4 = filenm4
	if filenm4 <> "" then 
		filename4 = filenm4.safeFileName	
		fileType4 = mid(filename4,inStrRev(filename4,".")+1)
		filename4 = company + "_" + sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_4." + fileType4
		save_path4 = path & "\" & filename4
	end if
	
	Set filenm5 = abc("att_file5")(1)
	filename5 = filenm5
	if filenm5 <> "" then 
		filename5 = filenm5.safeFileName	
		fileType5 = mid(filename5,inStrRev(filename5,".")+1)
		filename5 = company + "_" + sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_5." + fileType5
		save_path5 = path & "\" & filename5
	end if
	
	if (filenm1.length + filenm2.length + filenm3.length + filenm4.length + filenm4.length) > 1024*1024*8  then 
    	response.write "<script language=javascript>"
      	response.write "alert('파일 용량 8M를 넘으면 안됩니다.');"
		response.write "history.go(-1);"
      	response.write "</script>"
      	response.end
	End If

	if filenm1 <> "" then 
		filenm1.save save_path1
	end if
	if filenm2 <> "" then 
		filenm2.save save_path2
	end if
	if filenm3 <> "" then 
		filenm3.save save_path3
	end if
	if filenm4 <> "" then 
		filenm4.save save_path4
	end if
	if filenm5 <> "" then 
		filenm5.save save_path5
	end if
'2014-01-25 기존에 설치사진 첨부 (종료)


	if as_process = "완료" or as_process = "취소" then
		if write_date <> "" and write_cnt <> "" then
			w_date = formatdatetime(write_date,2)
			w_time = formatdatetime(write_date,4)
			w_sec = right(write_date,3)
			ww_date = w_date + " " + w_time + w_sec

			itft_type = as_type
			if as_type = "방문처리" then
				itft_type = "A/S방문"
			end if
			if as_type = "신규설치" or as_type = "신규설치공사" or as_type = "이전설치" or as_type = "이전설치공사" or as_type = "장비회수" then
				itft_type = "설치"
			end if

			sql = "Update nkp.as_acpt set dept ='"&dept&"', addr ='"&addr&"',mg_ce_id='"&mg_ce_id&"',mg_ce='"&mg_ce&"',as_memo='"&as_memo& _
			 "',request_date='"&request_date&"',request_time='"&request_time&"',visit_date='"&visit_date&"',visit_time='"&visit_time& _
			 "',arrival_date='"&visit_date&"',arrival_time='"&visit_time&"',into_reason='"&into_reason&"',as_process='"&as_process& _
			 "',as_type='"&itft_type&"',maker='"&maker&"',as_device='"&as_device&"',confirm_man='"&confirm_man&"',asets_no='"&asets_no& _
			 "',serial_no='"&serial_no &"',model_no='"&model_no& "',as_parts='"&as_parts&"',as_history='"&as_history&"',err_pc_sw='"&err01& _
			 "',err_pc_hw='"&err02&"',err_monitor='"&err03&"',err_printer='"&err04&"',err_network='"&err05&"',mod_date=now(),mod_id='"&mod_id& _
			 "',sms='"&sms_yn&"',before_process='"&as_process_old&"' where date_format(write_date,'%Y-%m-%d %H:%i:%s') = '"&ww_date& _
			 "' and write_cnt ="&write_cnt
'			response.write(sql)
			dbconn.execute(sql)
		end if		
		sql = "Update nkp.as_acpt set dept ='"&dept&"', addr ='"&addr&"',mg_ce_id='"&mg_ce_id&"',mg_ce='"&mg_ce&"',as_memo='"&as_memo& _
		"',request_date='"&request_date&"',request_time='"&request_time&"',visit_date='"&visit_date&"',visit_time='"&visit_time& _
		"',into_reason='"&into_reason&"',as_process='"&as_process&"',as_type='"&as_type&"',maker='"&maker&"',as_device='"&as_device& _
		"',asets_no='"&asets_no &"',serial_no='"&serial_no &"',model_no='"&model_no& "',as_parts='"&as_parts&"',as_history='"&as_history& _
		"',err_pc_sw='"&err01&"',err_pc_hw='"&err02&"',err_monitor='"&err03&"',err_printer='"&err04&"',err_network='"&err05& _
		"',err_server='"&err06&"',err_adapter='"&err07&"',err_etc='"&err09&"',dev_inst_cnt="&dev_inst_cnt&",ran_cnt="&ran_cnt& _
		",work_man_cnt="&work_man_cnt&",alba_cnt="&alba_cnt&",reside_place='"&reside_place&"',reside_company='"&reside_company& _
		"',reside='"&reside&"',team='"&team&"',mod_date=now(),mod_id='"&mod_id&"',sms='"&sms_yn&"',before_process='"&as_process_old& _
		"' where acpt_no = "&int(acpt_no)
		dbconn.execute(sql)
		
		if as_process = "완료" and work_man_cnt < 2 then
			sql = "delete from nkp.ce_work where acpt_no ="&int(acpt_no)
			dbconn.execute(sql)

			sql="insert into nkp.ce_work (acpt_no,mg_ce_id,work_id,work_date,as_type,company,emp_company,bonbu,saupbu,team,org_name"& _
			",reside_place"&",reside,reside_company,work_man_cnt,dev_inst_cnt,ran_cnt,alba_cnt,person_amt,reg_id,reg_date) values ('"&acpt_no& _
			"','"&mg_ce_id&"','2','"&visit_date&"','"&as_type&"','"&company&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")& _
			"','"&team&"','"&org_name&"','"&reside_place&"','"&reside&"','"&reside_company&"',"&work_man_cnt&","&dev_inst_cnt&","&ran_cnt& _
			","&alba_cnt&","&person_amt&",'"&user_id&"',now())"
			dbconn.execute(sql)
		  else
			sql = "Update nkp.ce_work set work_date ='"&visit_date&"' where work_id = '2' and acpt_no = "&int(acpt_no)
			dbconn.execute(sql)
		end if
		
	end if
	if as_process = "입고" then
		if write_date <> "" and write_cnt <> "" then
			w_date = formatdatetime(write_date,2)
			w_time = formatdatetime(write_date,4)
			w_sec = right(write_date,3)
			ww_date = w_date + " " + w_time + w_sec

			sql = "Update nkp.as_acpt set into_reason ='"&into_reason&"', as_process ='"&as_process&"',as_history='"&as_history& _
			"',mod_date=now(),mod_id='"&mod_id&"',sms='"&sms_yn&"',before_process='"&as_process_old& _
			"' where date_format(write_date,'%Y-%m-%d %H:%i:%s') = '"&ww_date&"' and write_cnt ="&write_cnt
			dbconn.execute(sql)
		end if		
		sql = "Update nkp.as_acpt set dept ='"&dept&"', addr ='"&addr&"',mg_ce_id='"&mg_ce_id&"',mg_ce='"&mg_ce&"',as_memo='"&as_memo& _
		"',request_date='"&request_date&"',request_time='"&request_time&"',into_reason='"&into_reason&"',as_process='"&as_process& _
		"',as_type='"&as_type&"',maker='"&maker&"',as_device='"&as_device&"',asets_no='"&asets_no &"',serial_no='"&serial_no& _
		"',model_no='"&model_no&"',in_date='"&in_date&"',in_replace='"&in_replace&"',as_parts='"&as_parts&"',as_history='"&as_history& _
		"',err_pc_sw='"&err01&"',err_pc_hw='"&err02&"',err_monitor='"&err03&"',err_printer='"&err04&"',err_network='"&err05& _
		"',err_server='"&err06&"',err_adapter='"&err07&"',err_etc='"&err09&"',reside_place='"&reside_place&"',reside='"&reside& _
		"',reside_company='"&reside_company&"',team='"&team&"',mod_date=now(),mod_id='"&mod_id&"',sms='"&sms_yn& _
		"',before_process='"&as_process_old&"' where acpt_no = "&int(acpt_no)
		dbconn.execute(sql)

		sql="insert into nkp.ce_work (acpt_no,mg_ce_id,work_id,work_date,as_type,company,emp_company,bonbu,saupbu,team,org_name,reside_place"& _
		",reside,reside_company,work_man_cnt,dev_inst_cnt,ran_cnt,alba_cnt,person_amt,reg_id,reg_date) values ('"&acpt_no&"','"&mg_ce_id& _
		"','3','"&in_date&"','"&as_type&"','"&company&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&team&"','"&org_name& _
		"','"&reside_place&"','"&reside&"','"&reside_company&"',"&work_man_cnt&","&dev_inst_cnt&","&ran_cnt&","&alba_cnt&","&person_amt& _
		",'"&user_id&"',now())"
		dbconn.execute(sql)
	end if
	if as_process = "접수" or as_process = "연기" then
		sql = "Update nkp.as_acpt set dept ='"&dept&"', addr ='"&addr&"',mg_ce_id='"&mg_ce_id&"',mg_ce='"&mg_ce&"',as_memo='"&as_memo& _
		"',request_date='"&request_date&"',request_time='"&request_time&"',into_reason='"&into_reason&"',as_process='"&as_process& _
		"',as_type='"&as_type&"',maker='"&maker&"',as_device='"&as_device&"',asets_no='"&asets_no &"',serial_no='"&serial_no& _
		"',model_no='"&model_no& "',as_parts='"&as_parts&"',as_history='"&as_history&"',err_pc_sw='"&err01&"',err_pc_hw='"&err02& _
		"',err_monitor='"&err03&"',err_printer='"&err04&"',err_network='"&err05&"',err_server='"&err06&"',err_adapter='"&err07& _
		"',err_etc='"&err09&"',reside_place='"&reside_place&"',reside_company='"&reside_company&"',reside='"&reside&"',team='"&team& _
		"',mod_date=now(),mod_id='"&mod_id&"',sms='"&sms_yn&"',before_process='"&as_process_old&"' where acpt_no = "&int(acpt_no)
		dbconn.execute(sql)
	end if

'2014-01-25 기존에 설치사진 첨부 (시작)
	if as_process = "완료" then
		if (filenm1 <> "") or (filenm2 <> "") or (filenm3 <> "") or (filenm4 <> "") or (filenm5 <> "") then 
            sql = "DELETE FROM nkp.att_file WHERE acpt_no = '"&acpt_no&"' "
            dbconn.execute(sql)

			sql = "insert into nkp.att_file (acpt_no,company,dept,sido,gugun,mg_ce_id,mg_ce,mg_group,visit_date,as_type,att_file1,att_file2,att_file3,att_file4,att_file5) values "
			sql = sql & "('"&acpt_no&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&mg_ce_id&"','"&mg_ce&"','"&mg_group&"','"&visit_date&"','"&as_type&"','"&filename1&"','"&filename2&"','"&filename3&"','"&filename4&"','"&filename5&"')"
			dbconn.execute(sql)
		end if
	end if
'2014-01-25 기존에 설치사진 첨부 (종료)

	sql="select * from nkp.juso_list where company='" & company & "' and dept = '" & dept & "'"
	set rs=dbconn.execute(sql)
	
	if	rs.eof or rs.bof then
		sql = "insert into nkp.juso_list (tel_ddd,tel_no1,tel_no2,company,dept,mg_group,sido,gugun,dong,addr,mg_ce_id,regi_date,regi_id,reside) values ('"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&company&"','"&dept&"','"&mg_group&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&mg_ce_id&"',now(),'"&regi_id&"','"&reside&"')"
		dbconn.execute(sql)
	  else
		sql = "update nkp.juso_list set addr='"&addr&"' where company='" & company & "' and dept = '" & dept & "'"
		dbconn.execute(sql)	  
	end if
' 변경 History 저장
	mod_pg = "결과등록"
	sql = "insert into nkp.as_mod (acpt_no,mod_date,mod_id,mod_name,mod_pg) values ('"&acpt_no&"',now(),'"&user_id&"','"&user_name&"','"&mod_pg&"')"
	dbconn.execute(sql)

	check_sw = "y"
	if be_pg = "as_list_user.asp" then
		url = "as_list.asp?page="&page&"&from_date="&from_date&"&to_date="&to_date&"&date_sw="&date_sw&"&process_sw="&process_sw&"&field_check="&field_check&"&field_view="&field_view&"&ck_sw="&check_sw&"&company="&condi_com
	  else
		url = "as_list_ce_user.asp?page="&page&"&view_sort="&view_sort&"&view_c="&view_c&"&ck_sw="&check_sw
	end if  
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + " 변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + " 변경되었습니다...."
	end if
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"alert('등록 완료 되었습니다....');"		
	response.write"location.replace('"&url&"');"
'	response.write"history.go(-2);"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
