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

	up_image = abc("page")

	path_name = "/image"
	path = Server.MapPath (path_name)

	Set filenm1 = abc("up_image")(1)
'	filename1 = filenm1
'	if filenm1 <> "" then 
	filename1 = filenm1.safeFileName	
	fileType1 = mid(filename1,inStrRev(filename1,".")+1)
'	filename1 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_1." + fileType1
'	end if
'	response.write(filename1)
	if filename1 = "nkp_popup.gif" or filename1 = "nkp_popup1.gif" then
		filename_ok = "Y"
	  else
	  	filename_ok = "N"
	end if

	if filename_ok = "N" then 
    	response.write "<script language=javascript>"
      	response.write "alert('파일명과 파일타입이 다릅니다');"
		response.write "history.go(-1);"
      	response.write "</script>"
      	response.end
	End If

	save_path1 = path & "\" & filename1
	filenm1.save save_path1
	
	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
