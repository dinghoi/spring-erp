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

	form_name = abc("form_name")
	company = abc("company")
	seq = abc("seq")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	path_nm = "D:\web\forms"

    Set fso=Server.CreateObject("Scripting.FileSystemObject")'
	if Not fso.FolderExists(path_nm) then
		path_nm = fso.CreateFolder(path_nm)
	end if
	Set fso = Nothing

	path_name = "/forms"
	path = Server.MapPath (path_name)

	Set filenm = abc("up_file")(1)
	filename = filenm
	if filenm <> "" then 
		filename = filenm.safeFileName	
		fileType = mid(filename,inStrRev(filename,".")+1)
		filename = company + "_" + form_name + "." + fileType
		save_path = path & "\" & filename
	end if
	
	if filenm <> "" then 
		filenm.save save_path
	end if

	sql = "select * from company_form where company = '"&company&"'"
	set rs = dbconn.execute(sql)
	if rs.eof or rs.bof then
		sql = "insert into company_form (company,form1,up_date1,up_id1) values ('"&company&"','"&filename&"',now(),'"&user_id&"')"
		response.write(sql)
		dbconn.execute(sql)
	  else
		if seq = 1 then
			sql = "Update company_form set form1 ='"&filename&"', up_date1 = now(), up_id1= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 2 then
			sql = "Update company_form set form2 ='"&filename&"', up_date2 = now(), up_id2= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 3 then
			sql = "Update company_form set form3 ='"&filename&"', up_date3 = now(), up_id3= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 4 then
			sql = "Update company_form set form4 ='"&filename&"', up_date4 = now(), up_id4= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 5 then
			sql = "Update company_form set form5 ='"&filename&"', up_date5 = now(), up_id5= '"&user_id&"' where company = '"&company&"'"
		end if
		dbconn.execute(sql)
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "변경되었습니다...."
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
