<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	dim code_tab(20)
	dim imsi_tab(20)
	dim del_tab(20)
	pummok_code = request("code_ary")+","
	del_check = request("del_ary")+","

	i=1
	j= 1
	jj=0
	k=0
	do until i=0
		i=0
		i=instr(j,del_check,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			del_tab(k)=""
	  	  else	  
			del_tab(k)=trim(mid(del_check,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	i=1
	j= 1
	jj=0
	k=0
	loop_cnt = 0
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
			loop_cnt = k
		end if
		j=i+1
		k=k+1
	loop

	j = 0
	for i = 0 to loop_cnt
		if del_tab(i) = "N" then
			j = j + 1
			code_tab(j) = imsi_tab(i)
		end if			
	next

	Set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	for i=1 to loop_cnt + 1
		if code_tab(i) <> "" then
			code_tab(i) = code_tab(i) + "/"
	
			j=instr(1,code_tab(i),"/")'
			code1 = trim(mid(code_tab(i),1,j-1))
	
			j1=instr(j+1,code_tab(i),"/")'
			code2 = trim(mid(code_tab(i),j+1,j1-(j+1)))
	
			j2=instr(j1+1,code_tab(i),"/")'
			code3 = trim(mid(code_tab(i),j1+1,j2-(j1+1)))
	
			j3=instr(j2+1,code_tab(i),"/")'
			code4 = trim(mid(code_tab(i),j2+1,j3-(j2+1)))
	
			j4=instr(j3+1,code_tab(i),"/")'
			code5 = trim(mid(code_tab(i),j3+1,j4-(j3+1)))
	
	'		Sql="select * from etc_code where etc_type = '51' and etc_code='"+code_tab(i)+"'"
			if code1 > "5100" and code1 < "5200" then
				Sql="select * from etc_code where etc_code='"+code1+"'"
			  else
				sql = "select goods_code as etc_code, goods_type as type_name, goods_gubun as etc_name, goods_standard as group_name from met_goods_code where goods_code = '"&code1&"' and goods_type = '상품' order by etc_name"
			end if	

			Set rs=DbConn.Execute(Sql)
			if rs.eof or rs.bof then
				response.write"<script language=javascript>"
				response.write"opener.document.frm.srv_type"&i&".value = '';"
				response.write"opener.document.frm.pummok_code"&i&".value = '';"
				response.write"opener.document.frm.pummok"&i&".value = '';"
				response.write"opener.document.frm.standard"&i&".value = '';"
				response.write"opener.document.frm.qty"&i&".value = '0';"
				response.write"opener.document.frm.buy_cost"&i&".value = '0';"
				response.write"opener.document.frm.sales_cost"&i&".value = '0';"
				response.write"opener.document.frm.margin_cost"&i&".value = '0';"
				response.write"opener.document.frm.sales_tot"&i&".value = '0';"
				response.write"opener.document.frm.margin_tot"&i&".value = '0';"
				response.write"opener.document.frm.del_check"&i&".checked = false;"
				response.write"opener.document.getElementById('pummok_list"&i&"').style.display = 'none';"
				response.write"</script>"
			  else
				response.write"<script language=javascript>"
				response.write"opener.document.frm.srv_type"&i&".value = '"&rs("type_name")&"';"
				response.write"opener.document.frm.pummok_code"&i&".value = '"&rs("etc_code")&"';"
				response.write"opener.document.frm.pummok"&i&".value = '"&rs("etc_name")&"';"
				response.write"opener.document.frm.standard"&i&".value = '"&code2&"';"
				response.write"opener.document.frm.qty"&i&".value = '"&code3&"';"
				response.write"opener.document.frm.buy_cost"&i&".value = '"&code4&"';"
				response.write"opener.document.frm.sales_cost"&i&".value = '"&code5&"';"
				response.write"opener.document.frm.del_check"&i&".checked = false;"
				response.write"opener.document.getElementById('pummok_list"&i&"').style.display = '';"
				response.write"</script>"
			end if

		end if
	next
	response.write"<script language=javascript>"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing				
%>

