<!--#include virtual="/common/inc_top.asp"-->
<%
'===================================================
'### Request & Params
'===================================================
Dim uploadForm, filenm, up_image
Dim path_name, path, filenm1, save_path1
Dim filename1, fileType1, filename_ok

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")
uploadForm.AbsolutePath = True
uploadForm.Overwrite = true
uploadForm.MaxUploadSize = 1024*1024*50

up_image = uploadForm("page")

path_name = "/image"
path = Server.MapPath(path_name)

Set filenm1 = uploadForm("up_image")(1)

'	filename1 = filenm1
'	if filenm1 <> "" then
filename1 = filenm1.safeFileName
fileType1 = Mid(filename1,inStrRev(filename1,".") + 1)
'	filename1 = company + "_" + o_sido + "_" + cstr(mid(visit_date,3,2)) + cstr(mid(visit_date,6,2)) + cstr(mid(visit_date,9,2)) + "_" + cstr(acpt_no) + "_1." + fileType1
'	end if

If filename1 = "nkp_popup.gif" Or filename1 = "nkp_popup1.gif" Or filename1 = "nkp_popup1.png" Then
	filename_ok = "Y"
Else
	filename_ok = "N"
End If

If filename_ok = "N" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('파일명과 파일타입이 다릅니다');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If

save_path1 = path & "\" & filename1
filenm1.save save_path1

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('등록 완료 되었습니다.');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close() ;"
Response.Write "</script>"
Response.End
%>
