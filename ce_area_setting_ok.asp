<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim ary_ce_id(50)
	dim ary_gugun(50)
	
	sido = request("sido")
	gugun = request("gugun")+","
	mod_ce_id = request("mod_ce_id")+","
	
	i=1
	j= 1
	jj=0
	k=0
	do until i=0
		i=0
		i=instr(j,gugun,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			ary_gugun(k)=""
	  	  else	  
			ary_gugun(k)=trim(mid(gugun,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	i=1
	j= 1
	jj=0	
	k=0
	do until i=0
		i=0
		i=instr(j,mod_ce_id,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			ary_ce_id(k)=""
	  	  else	  
			ary_ce_id(k)=trim(mid(mod_ce_id,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	set dbconn = server.CreateObject("adodb.connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

	mod_cnt = 0

	for i = 0 to k-1 step 1

		sql = "select * from memb where user_id='" + ary_ce_id(i) + "'"
		Set rs=DbConn.Execute(sql)
		if not rs.eof then
			c_grade = rs("grade")
			rs.close()
			if ary_ce_id(i) <> "" then
				sql = "update ce_area set mg_ce_id='"&ary_ce_id(i)&"', mod_date=now() , mod_id='"&user_id&"' where mg_group='"+mg_group+"' and sido='" + sido + "' and gugun='" + ary_gugun(i) + "'"
				dbconn.execute(sql)

				sql = "update area_mg set mg_ce_id='"&ary_ce_id(i)&"', mod_date=now() , mod_id='"&user_id&"' where mg_group='"+mg_group+"' and sido='" + sido + "' and gugun='" + ary_gugun(i) + "'"
				dbconn.execute(sql)

				sql = "update juso_list set mg_ce_id='"&ary_ce_id(i)&"', regi_date=now() , regi_id='"&user_id&"' where mg_group='"+mg_group+"' and reside = '0' and sido='" + sido + "' and gugun='" + ary_gugun(i) + "'"
				dbconn.execute(sql)

				mod_cnt = mod_cnt + 1
			end if
		end if
	next
	if	mod_cnt = 0 then
		response.write"<script language=javascript>"
		response.write"alert('변경된 내역이 없읍니다.');"
		response.write"location.replace('ce_area_setting.asp?sido=" & sido & "');"
		response.write"</script>"	
		dbconn.Close()
		Set dbconn = Nothing
	  else
		msg = cstr(mod_cnt) + " 건 변경되었습니다."
		response.write"<script language=javascript>"
		response.write"alert('변경 완료되었습니다.');"
		response.write"location.replace('ce_area_setting.asp?sido=" & sido & "');"
		response.write"</script>"	

		dbconn.Close()
		Set dbconn = Nothing

	end if

%>

