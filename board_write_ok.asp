<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	dim abc,filenm
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
	
	fileType = mid(filename,inStrRev(filename,".")+1)

	save_path = path & "\" & filename
		
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open DbConnect

	if filenm.length > 1024*1024*8  then 
    	response.write "<script language=javascript>"
      	response.write "alert('파일 용량 8M를 넘으면 안됩니다.');"
		response.write "history.go(-1);"
      	response.write "</script>"
      	response.end
	End If

	url = "nkp_main.asp?board_gubun="&board_gubun&"&page="&page&"&condi="&condi&"&condi_value="&condi_value&"&ck_sw=y"					

	if u_type = "U" then

		Sql="select * from board where board_seq="&board_seq
		Set Rs=dbconn.execute(Sql)
	
		if	rs("pass") <> pass then
			response.write"<script language=javascript>"
			response.write"alert('입력하신 비밀번호가 틀립니다.');"
			response.write"history.go(-1);"
			response.write"</script>"
		  Else
			if filenm <> "" then 
				filenm.save save_path
				sql = "Update board set board_title ='"&board_title&"', board_body='"&board_body&"', mod_date=now(), att_file='"&filename&"' where board_seq = "&board_seq
			  Else
				sql = "Update board set board_title ='"&board_title&"', board_body='"&board_body&"', mod_date=now()  where board_seq = "&board_seq
			end if 				
			dbconn.execute(sql)
			response.write"<script language=javascript>"
			response.write"alert('등록 완료 되었습니다....');"		
'			response.write"location.replace('"&url&"');"
			response.write"parent.opener.location.reload();"
			response.write"self.close() ;"
			response.write"</script>"		
			Response.End
		end if			
	  else			
		if filenm <> "" then 
			filenm.save save_path
			sql = "insert into board (board_gubun,reg_id,reg_name,board_title,ed_sw,board_body,pass,reg_date,read_cnt,att_file)  values ('"&board_gubun&"','"&user_id&"','"&user_name&"','"&board_title&"','"&ed_sw&"','"&board_body&"','"&pass&"', now(),0,'"&filename&"')"
		  Else
			sql = "insert into board (board_gubun,reg_id,reg_name,board_title,ed_sw,board_body,pass,reg_date,read_cnt) values ('"&board_gubun&"','"&user_id&"','"&user_name&"','"&board_title&"','"&ed_sw&"','"&board_body&"','"&pass&"', now(),0)"
  		end if 	
		dbconn.execute(sql)
		
		response.write"<script language=javascript>"
		response.write"alert('등록 완료 되었습니다....');"		
'		response.write"location.replace('"&url&"');"
		response.write"parent.opener.location.reload('"&url&"');"
		response.write"self.close() ;"
		response.write"</script>"		
		Response.End
	end if

	dbconn.close()
	Set dbconn = nothing
%>
