<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	dim code_tab(41)
	dim imsi_tab(41)
	code_ary = request.form("code_ary")+","
	pummok_code = request.form("sel_check")+","
	pummok_code = code_ary + pummok_code
	
	goods_type = Request.form("goods_type1") 

'response.write(pummok_code)

				
	i=1
	j= 1
	jj=0
	k=0

    do until i=0
		i=0
		i=instr(j,pummok_code,",")'
'response.write(i)
'response.write("/")
'response.End()		
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			imsi_tab(k)=""
	  	  else	  
'response.write(k)
'response.write("k")
'response.End()
			
			imsi_tab(k)=trim(mid(pummok_code,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

'response.End()	
	
	k = 0
	for i = 0 to 41
		if imsi_tab(i) <> "" and imsi_tab(i) <> "/"  then
			code_tab(k) = imsi_tab(i)
			k = k + 1
		end if
	next

	Set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	
	j = 0
	for i=0 to 41
'		if code_tab(i) = "" or code_tab(i) = "/" then
		if code_tab(i) = "" then
			exit for
		end if
		
		code2 = ""
		code_tab(i) = code_tab(i) + "/"
		k=instr(1,code_tab(i),"/")'
		code1 = trim(mid(code_tab(i),1,k-1))

'		k1=instr(k+1,code_tab(i),"/")'
'		if k1 = 0 then
'			code2 = ""
'		  else
'			code2 = trim(mid(code_tab(i),k+1,k1-(k+1)))
'		end if
		
		j = i + 1		
'response.write(code1)
		Sql="select * from met_goods_code where goods_code='"+code1+"'"
'		response.write(j)
'		response.write("==")
'		response.write(sql)
'		response.write("**")
		Set rs=DbConn.Execute(Sql)
		
				
		response.write"<script language=javascript>"
		response.write"opener.document.frm.srv_type"&j&".value = '"&goods_type&"';"
		response.write"opener.document.frm.goods_gubun"&j&".value = '"&rs("goods_gubun")&"';"
		response.write"opener.document.frm.goods_code"&j&".value = '"&rs("goods_code")&"';"
		response.write"opener.document.frm.goods_name"&j&".value = '"&rs("goods_name")&"';"
		response.write"opener.document.frm.part_number"&j&".value = '"&rs("part_number")&"';"
		response.write"opener.document.getElementById('pummok_list"&j&"').style.display = '';"

		response.write"</script>"
	next
	response.write"<script language=javascript>"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>

