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

	acpt_no = int(abc("acpt_no"))
	sido = abc("sido")
	company = abc("company")
	visit_date = abc("visit_date")
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

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

	sql = "Update att_file set att_file1='"&filename1&"',att_file2='"&filename2&"',att_file3='"&filename3&"',att_file4='"&filename4& _
	"',att_file5='"&filename5&"',mod_id='"&user_id&"',mod_date=now() where acpt_no = "&int(acpt_no)
	dbconn.execute(sql)
		
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
