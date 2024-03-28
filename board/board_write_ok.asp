<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim abc, filenm, board_seq, board_gubun, board_title
Dim pass, end_sw, boar_body, v_att_file, u_type
Dim condi, condi_value, page, ck_sw, fileType
Dim save_path, ed_sw, board_body, filename
Dim path, url, rs

Set abc = Server.CreateObject("ABCUpload4.XForm")
abc.AbsolutePath = True
abc.Overwrite = true
abc.MaxUploadSize = 1024*1024*50

board_seq = abc("board_seq")
board_gubun = abc("board_gubun")
board_title = abc("board_title")
pass = abc("pass")
ed_sw = "Y"
board_body = abc("board_body")
board_body = Replace(board_body,"'","&quot;")
v_att_file= abc("v_att_file")
u_type = abc("u_type")
condi = abc("condi")
condi_value = abc("condi_value")
page = abc("page")
ck_sw = abc("ck_sw")

Set filenm = abc("att_file")(1)

path = Server.MapPath ("/nkp_upload")
filename = filenm.safeFileName

fileType = Mid(filename,inStrRev(filename,".") + 1)

save_path = path & "\" & filename

If filenm.length > 1024*1024*8  Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('파일 용량 8M를 넘으면 안됩니다.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If

url = "/main/nkp_main.asp?board_gubun="&board_gubun&"&page="&page&"&condi="&condi&"&condi_value="&condi_value&"&ck_sw=y"

If u_type = "U" Then
'	Response.write "update"
'	Response.end
	'Sql = "select * from board where board_seq="&board_seq
	objBuilder.Append "SELECT pass FROM board WHERE board_seq = " & board_seq

	Set rs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs("pass") <> pass Then
		Response.Write "<script type='text/javascript'>"
		Response.Write "	alert('입력하신 비밀번호가 틀립니다.');"
		Response.Write "	history.go(-1);"
		Response.Write "</script>"
	Else
		If filenm <> "" Then
			filenm.save save_path

			'sql = "Update board set board_title ='"&board_title&"', board_body='"&board_body&"', mod_date=now(), att_file='"&filename&"' where board_seq = "&board_seq
			objBuilder.Append "UPDATE board SET "
			objBuilder.Append "board_title = '"&board_title&"', board_body = '"&board_body&"', "
			objBuilder.Append "mod_date = NOW(), att_file = '"&filename&"' "
			objBuilder.Append "WHERE board_seq = " & board_seq
		Else
			'sql = "Update board set board_title ='"&board_title&"', board_body='"&board_body&"', mod_date=now()  where board_seq = "&board_seq
			objBuilder.Append "UPDATE board SET "
			objBuilder.Append "board_title = '"&board_title&"', board_body = '"&board_body&"', mod_date = NOW() "
			objBuilder.Append "WHERE board_seq = "&board_seq
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		Response.Write "<script type='text/javascript'>"
		Response.Write "	alert('등록 완료 되었습니다.');"
		'Response.Write "	location.replace('"&url&"');"
		Response.Write "	parent.opener.location.reload();"
		Response.Write "	self.close();"
		Response.Write "</script>"
		Response.End
	End If

	rs.Close() : Set rs = Nothing
Else
	If filenm <> "" Then
		filenm.save save_path

		'sql = "insert into board (board_gubun,reg_id,reg_name,board_title,ed_sw,board_body,pass,reg_date,read_cnt,att_file)  values ('"&board_gubun&"','"&user_id&"','"&user_name&"','"&board_title&"','"&ed_sw&"','"&board_body&"','"&pass&"', now(),0,'"&filename&"')"
		objBuilder.Append "INSERT INTO board(board_gubun, reg_id, reg_name, board_title, ed_sw, "
		objBuilder.Append "board_body, pass, reg_date, read_cnt, att_file)VALUES("
		objBuilder.Append "'"&board_gubun&"','"&user_id&"','"&user_name&"','"&board_title&"', '"&ed_sw&"', "
		objBuilder.Append "'"&board_body&"','"&pass&"', NOW(), 0, '"&filename&"')"
	Else
		'sql = "insert into board (board_gubun,reg_id,reg_name,board_title,ed_sw,board_body,pass,reg_date,read_cnt) values ('"&board_gubun&"','"&user_id&"','"&user_name&"','"&board_title&"','"&ed_sw&"','"&board_body&"','"&pass&"', now(),0)"
		objBuilder.Append "INSERT INTO board (board_gubun,reg_id,reg_name,board_title,ed_sw, "
		objBuilder.Append "board_body,pass,reg_date,read_cnt)VALUES("
		objBuilder.Append "'"&board_gubun&"','"&user_id&"','"&user_name&"','"&board_title&"','"&ed_sw&"', "
		objBuilder.Append "'"&board_body&"','"&pass&"', NOW(), 0)"

	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('등록 완료 되었습니다.');"
	'Response.Write "	location.replace('"&url&"');"
	Response.Write "	parent.opener.location.reload('"&url&"');"
	Response.Write "	self.close();"
	Response.Write "</script>"
	Response.End
End If

DBConn.Close() : Set DBConn = Nothing
%>
