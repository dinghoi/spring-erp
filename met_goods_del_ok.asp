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
	
	goods_type = Request("goods_type")

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

    buy_tot_cost = 0
	for i=1 to loop_cnt + 1
'      if code_tab(i) <> "" then
'		if code_tab(i) = "" then
'			exit for
'		end if
		
		Sql="select * from met_goods_code where goods_code='"&code_tab(i)&"'"
		Set rs=DbConn.Execute(Sql)
		if rs.eof or rs.bof then
			response.write"<script language=javascript>"
			response.write"opener.document.frm.srv_type"&i&".value = '';"
			response.write"opener.document.frm.goods_gubun"&i&".value = '';"
			response.write"opener.document.frm.goods_code"&i&".value = '';"
			response.write"opener.document.frm.goods_name"&i&".value = '';"
			response.write"opener.document.frm.goods_standard"&i&".value = '';"
			response.write"opener.document.frm.qty"&i&".value = '0';"
			response.write"opener.document.frm.buy_cost"&i&".value = '0';"
			response.write"opener.document.frm.buy_tot"&i&".value = '0';"
			response.write"opener.document.frm.del_check"&i&".checked = false;"
			response.write"opener.document.getElementById('pummok_list"&i&"').style.display = 'none';"
			response.write"</script>"
		  else
			response.write"<script language=javascript>"
			response.write"opener.document.frm.srv_type"&i&".value = '"&goods_type&"';"
			response.write"opener.document.frm.goods_gubun"&i&".value = '"&rs("goods_gubun")&"';"
			response.write"opener.document.frm.goods_code"&i&".value = '"&rs("goods_code")&"';"
			response.write"opener.document.frm.goods_name"&i&".value = '"&rs("goods_name")&"';"
			response.write"opener.document.frm.goods_standard"&i&".value = '"&rs("goods_standard")&"';"
			response.write"opener.document.frm.del_check"&i&".checked = false;"
			response.write"opener.document.getElementById('pummok_list"&i&"').style.display = '';"
			response.write"</script>"
		end if
		
'	   end if
	next

	response.write"<script language=javascript>"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing				
%>

