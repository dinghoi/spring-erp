<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim ary_holiday(20)
	holiday = request.form("del_ck")+","
		
	i=1
	j= 1
	jj=0
	k=0
	do until i=0
		i=0
		i=instr(j,holiday,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			ary_holiday(k)=""
	  	  else	  
			ary_holiday(k)=trim(mid(holiday,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	Set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	j = 0
	for i=0 to 20
		if ary_holiday(i) = "" then
			exit for
		end if

		Sql="delete from holiday where holiday='"+ary_holiday(i)+"'"
		dbconn.execute(sql)
		j = j + 1
	next
	url = "holi_mg.asp"
	del_msg = cstr(j) + "건 삭제 되었습니다."
	response.write"<script language=javascript>"
	response.write"alert('"&del_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"		

	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>

