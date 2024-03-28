<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	dim code_tab(20)
	dim imsi_tab(20)
	slip_id = request.form("slip_id")
	srv_type = request.form("srv_type")
	code_ary = request.form("code_ary")+","
	pummok_code = request.form("sel_check")+","
'	pummok_code = pummok_code + code_ary
	pummok_code = code_ary + pummok_code

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
			imsi_tab(k)=""
	  	  else	  
			imsi_tab(k)=trim(mid(pummok_code,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop
	k = 0
	for i = 0 to 20
		if imsi_tab(i) <> "" and imsi_tab(i) <> "/"  then
			code_tab(k) = imsi_tab(i)
			k = k + 1
		end if
	next

	Set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	j = 0
	for i=0 to 20
'		if code_tab(i) = "" or code_tab(i) = "/" then
		if code_tab(i) = "" then
			exit for
		end if

		code2 = ""
		code_tab(i) = code_tab(i) + "/"
		k=instr(1,code_tab(i),"/")'
		code1 = trim(mid(code_tab(i),1,k-1))

		k1=instr(k+1,code_tab(i),"/")'
		if k1 = 0 then
			code2 = ""
		  else
			code2 = trim(mid(code_tab(i),k+1,k1-(k+1)))
		end if
		
		j = i + 1		
		if code1 > "5100" and code1 < "5200" then
			Sql="select * from etc_code where etc_code='"+code1+"'"
		  else
			sql = "select stock_goods_code as etc_code, stock_goods_type as type_name, stock_goods_gubun as etc_name, stock_goods_standard as group_name from met_stock_gmaster where stock_goods_code = '"&code1&"' and stock_goods_type = '상품' order by etc_name"
		end if
		response.write(j)
		response.write("==")
		response.write(sql)
		response.write("**")
		Set rs=DbConn.Execute(Sql)
		if code2 = "" then
			if rs("group_name") = "" or isnull(rs("group_name")) then
				standard_view = rs("group_name")
			  else			
				standard_view = Replace(rs("group_name"),","," ")
		  	end if
		  else
		  	standard_view = code2
		end if

		response.write"<script language=javascript>"
		response.write"opener.document.frm.srv_type"&j&".value = '"&rs("type_name")&"';"
		response.write"opener.document.frm.pummok_code"&j&".value = '"&rs("etc_code")&"';"
		response.write"opener.document.frm.pummok"&j&".value = '"&rs("etc_name")&"';"
		response.write"opener.document.frm.standard"&j&".value = '"&standard_view&"';"
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

