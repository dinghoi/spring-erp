<%@LANGUAGE="VBSCRIPT"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'	on Error resume next

	dim date_tab(10)
	dim seq_tab(10)
	dim confirm_tab(10)
	
	slip_month = request.form("slip_month")
	view_condi = request.form("view_condi")
	condi = request.form("condi")
	acpt_confirm = request.form("acpt_confirm")
	page = request.form("page")	
	tot_seq = int(request.form("tot_seq"))
	slip_date = cstr(request.form("slip_date"))+","
	slip_seq = request.form("slip_seq")+","
	confirm = request.form("confirm_yn")+","
		
	i=1
	j= 1
	jj=0
	k=1
	do until i=0
		i=0
		i=instr(j,confirm,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			confirm_tab(k)=""
	  	  else	  
			confirm_tab(k)=trim(mid(confirm,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	i=1
	j= 1
	jj=0
	k=1
	do until i=0
		i=0
		i=instr(j,slip_date,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			date_tab(k)=""
	  	  else	  
			date_tab(k)=trim(mid(slip_date,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	i=1
	j= 1
	jj=0
	k=1
	do until i=0
		i=0
		i=instr(j,slip_seq,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			seq_tab(k)=""
	  	  else	  
			seq_tab(k)=trim(mid(slip_seq,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	dbconn.BeginTrans

	for i = 1 to 10
		if confirm_tab(i) = "" then
			exit for
		end if			
		j = int(confirm_tab(i))

		sql = "update general_cost set confirm_yn='Y',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where slip_date='"&date_tab(i)&"' and slip_seq = '"&seq_tab(i)&"'"
		dbconn.execute(sql)	  
	next

	url = "general_cost_check.asp?page="&page&"&slip_month="&slip_month&"&confirm="&acpt_confirm&"&ck_sw=y&view_condi="&view_condi&"&condi="&condi

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>

