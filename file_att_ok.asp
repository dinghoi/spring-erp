<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%

	dim abc,filenm

	Set abc = Server.CreateObject("ABCUpload4.XForm")

	abc.AbsolutePath = True
	abc.Overwrite = true

	Set filenm = abc("fileData")(1)
	
	path = Server.MapPath ("/kwon_upload")
	filename = filenm.safeFileName
	filename = "12345"
	filetype = filenm.FileType
	save_path = path & "\" & filename & "." & filetype
	response.write(save_path)
	if filenm.length < 512000  then 
		filenm.save save_path
'	  	If filename <> "" Then
'			if filenm <> "" then 
'			sql = "insert into dbo.k1_board (gubun,id,name,title,ed_sw,body,pass,w_date,cnt,att_file,mg_group) " & _
'	      	"values ('"&gubun&"','"&id&"','"&sname&"','"&title&"','"&ed_sw&"','"&body&"','"&pass&"', getdate(),0,'"&filename&"','"&mg_group&"')"
'		Else
'			sql = "insert into dbo.k1_board (gubun,id,name,title,ed_sw,body,pass,w_date,cnt,mg_group) " & _
'	      	"values ('"&gubun&"','"&id&"','"&sname&"','"&title&"','"&ed_sw&"','"&body&"','"&pass&"', getdate(),0,'"&mg_group&"')"						
'  		end if 
	
'		set dbconn = server.CreateObject("adodb.connection")
'		dbconn.open DbConnect
'		dbconn.execute(sql)
		
		response.write"<script language=javascript>"
		response.write"alert('등록 완료 되었습니다....');"		
		response.write"location.replace('file_att.asp');"
		response.write"</script>"		
		Response.End

	Else
    	response.write "<script language=javascript>"
      	response.write "alert('파일 용량 500K를 넘으면 안됩니다.');"
		response.write "history.back();"
      	response.write "</script>"
      	response.end
	End If

%>
