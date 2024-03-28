<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim code_tab(20)
	code_ary = request.form("code_ary")+","
	response.write(code_ary)
	response.write("===")
	pummok_code = request.form("sel_check")+","
	response.write(pummok_code)
	response.write("===")
	pummok_code = pummok_code + code_ary
	response.write(pummok_code)
	response.write("===")
				
	i=1
	j= 1
	jj=0
	k=0
	do until i=0
		i=0
		i=instr(j,pummok_code,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			code_tab(k)=""
	  	  else	  
			code_tab(k)=trim(mid(pummok_code,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	Set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	j = 0
	for i=0 to 20
		if code_tab(i) = "" then
			exit for
		end if

		j = i + 1
		

		Sql="select * from pummok_code where pummok_code='"+code_tab(i)+"'"
		Set rs=DbConn.Execute(Sql)
		response.write"<script language=javascript>"
		response.write"opener.document.frm.srv_type"&j&".value = '"&rs("srv_type")&"';"
		response.write"opener.document.frm.pummok_code"&j&".value = '"&rs("pummok_code")&"';"
		response.write"opener.document.frm.pummok"&j&".value = '"&rs("pummok_name")&"';"
		response.write"opener.document.frm.standard"&j&".value = '"&rs("standard")&"';"
		response.write"opener.document.getElementById('pummok_list"&j&"').style.display = '';"
		response.write"</script>"

	next
	response.write"<script language=javascript>"
'	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>

