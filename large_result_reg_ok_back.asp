<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

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

	acpt_no = int(abc("acpt_no"))
	company = abc("company")
	dept = abc("dept")
	as_type = abc("as_type")
	o_sido = abc("o_sido")
	o_gugun = abc("o_gugun")
	o_dong = abc("o_dong")
	o_addr = abc("o_addr")
	juso_mod_ck = abc("juso_mod_ck")
	if	juso_mod_ck = "1" then		
		sido = abc("sido")
		gugun = abc("gugun")
		dong = abc("dong")
		addr = abc("addr")
	  else
	  	juso_mod_ck = "0"
	end if

	visit_date = abc("visit_date")
	visit_hh = abc("visit_hh")
	visit_mm = abc("visit_mm")
	visit_time = cstr(visit_hh) + cstr(visit_mm)
  	dev_inst_cnt = int(abc("dev_inst_cnt1"))
	ran_cnt = int(abc("ran_cnt"))
	work_man_cnt = int(abc("work_man_cnt"))
	alba_cnt = int(abc("alba_cnt"))
	person_amt = dev_inst_cnt + ran_cnt
	mg_ce_id = user_id
	be_pg = abc("be_pg")
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select * from memb where user_id = '" + user_id + "'"
	set rs=dbconn.execute(sql)

	if	rs.eof or rs.bof then
		team = "error"
		org_name = "error"
	  else
		team = rs("team")
		org_name = rs("org_name")
	end if

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
		filename1 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_1." + fileType1
		save_path1 = path & "\" & filename1
	end if
	
	Set filenm2 = abc("att_file2")(1)
	filename2 = filenm2
	if filenm2 <> "" then 
		filename2 = filenm2.safeFileName	
		fileType2 = mid(filename2,inStrRev(filename2,".")+1)
		filename2 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_2." + fileType2
		save_path2 = path & "\" & filename2
	end if
	
	Set filenm3 = abc("att_file3")(1)
	filename3 = filenm3
	if filenm3 <> "" then 
		filename3 = filenm3.safeFileName	
		fileType3 = mid(filename3,inStrRev(filename3,".")+1)
		filename3 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_3." + fileType3
		save_path3 = path & "\" & filename3
	end if
	
	Set filenm4 = abc("att_file4")(1)
	filename4 = filenm4
	if filenm4 <> "" then 
		filename4 = filenm4.safeFileName	
		fileType4 = mid(filename4,inStrRev(filename4,".")+1)
		filename4 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_4." + fileType4
		save_path4 = path & "\" & filename4
	end if
	
	Set filenm5 = abc("att_file5")(1)
	filename5 = filenm5
	if filenm5 <> "" then 
		filename5 = filenm5.safeFileName	
		fileType5 = mid(filename5,inStrRev(filename5,".")+1)
		filename5 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_5." + fileType5
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

	if	juso_mod_ck = "1" then		
		sql = "Update as_acpt set sido ='"+sido+"',gugun ='"+gugun+"',dong ='"+dong+"',addr ='"+addr+"',mg_ce_id='"+user_id
		sql = sql + "',mg_ce='"+user_name+"',visit_date='"+visit_date+"',visit_time='"+visit_time+"',as_process='완료',as_type='"+as_type
		sql = sql + "',as_history='"+as_type+"',dev_inst_cnt="&dev_inst_cnt&",ran_cnt="&ran_cnt
		sql = sql + ",work_man_cnt="&work_man_cnt&",alba_cnt="&alba_cnt&",org_name='"+org_name+"',team='"+team
		sql = sql + "',mod_date=now(),mod_id='"+user_id+"',before_process='접수' where acpt_no = "&int(acpt_no)
		dbconn.execute(sql)
		sql="insert into old_juso (acpt_no,sido,gugun,dong,addr) values ('"&acpt_no&"','"&o_sido&"','"&o_gugun&"','"&o_dong&"','"&o_addr&"')"
		dbconn.execute(sql)
	  else
		sql = "Update as_acpt set mg_ce_id='"+user_id+"',mg_ce='"+user_name+"',visit_date='"+visit_date+"',visit_time='"+visit_time
		sql = sql +"',as_process='완료',as_type='"+as_type+"',as_history='"+as_type+"',dev_inst_cnt="&dev_inst_cnt&",ran_cnt="&ran_cnt
		sql = sql + ",work_man_cnt="&work_man_cnt&",alba_cnt="&alba_cnt&",org_name='"+org_name+"',team='"+team
		sql = sql + "',mod_date=now(),mod_id='"+user_id+"',before_process='접수' where acpt_no = "&int(acpt_no)
		dbconn.execute(sql)
	end if

		
	if work_man_cnt < 2 then
		sql="insert into ce_work (acpt_no,mg_ce_id,work_id,work_date,as_type,company,emp_company,bonbu,saupbu,team,org_name,reside,work_man_cnt,dev_inst_cnt,ran_cnt,alba_cnt,person_amt,reg_id,reg_date) values ('"&acpt_no&"','"&user_id&"','2','"&visit_date& _
		"','"&as_type&"','"&company&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")& _
		"','"&rs("reside")&"',"&work_man_cnt&","&dev_inst_cnt&","&ran_cnt&","&alba_cnt&","&person_amt&",'"&user_id&"',now())"
		dbconn.execute(sql)
	  else
		sql = "Update ce_work set work_date ='"&visit_date&"' where work_id = '2' and acpt_no = "&int(acpt_no)
		dbconn.execute(sql)
	end if
		
	sql = "insert into att_file (acpt_no,company,dept,sido,gugun,mg_ce_id,mg_ce,mg_group,visit_date,as_type,att_file1,att_file2,att_file3,att_file4,att_file5) values "
	sql = sql + "('"&acpt_no&"','"&company&"','"&dept&"','"&sido&"','"&gugun&"','"&user_id&"','"&user_name&"','"&mg_group&"','"&visit_date&"','"&as_type&"','"&filename1&"','"&filename2&"','"&filename3&"','"&filename4&"','"&filename5&"')"
	dbconn.execute(sql)

	check_sw = "y"
	if be_pg = "as_list.asp" then
		url = "as_list.asp?page="+page+"&from_date="+from_date+"&to_date="+to_date+"&date_sw="+date_sw+"&process_sw="+process_sw+"&field_check="+field_check+"&field_view="+field_view+"&ck_sw="+check_sw+"&company="+condi_com
	  else
		url = "as_list_ce.asp?page="+page+"&view_sort="+view_sort+"&view_c="+view_c+"&ck_sw="+check_sw
	end if  
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "변경되었습니다...."
	end if
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"alert('등록 완료 되었습니다....');"		
'	response.write"location.replace('"&url&"');"
'	response.write"history.go(-2);"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
