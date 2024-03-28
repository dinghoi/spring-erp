<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/asmg_dbcon.asp" -->
<%
	dim ary_ce_id(20)
	mg_ce_id = request.form("del_ck")+","
	page = request("page")
	page_cnt = request("page_cnt")
	view_condi = request("view_condi")
	condi = request("condi")
		
	i=1
	j= 1
	jj=0
	k=0
	do until i=0
		i=0
		i=instr(j,mg_ce_id,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			ary_ce_id(k)=""
	  	  else	  
			ary_ce_id(k)=trim(mid(mg_ce_id,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	Set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	j = 0
	for i=0 to 20
		if ary_ce_id(i) = "" then
			exit for
		end if

		Sql="delete from memb where user_id='"+ary_ce_id(i)+"'"
'		response.write(sql)
		dbconn.execute(sql)
		j = j + 1
	next
	url = "ce_mg_list.asp?page=" + page + "&page_cnt=" + page_cnt + "&ck_sw= y&view_condi="+view_condi+"&condi="+ condi
	del_msg = cstr(j) + "건 삭제 되었습니다."
	response.write"<script language=javascript>"
	response.write"alert('"&del_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"		

	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>

