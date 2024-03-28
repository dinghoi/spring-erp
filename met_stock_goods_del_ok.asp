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
    stock_code = Request("stock_code") 
	chul_date = Request("chul_date") 
	chul_seq = Request("chul_seq") 

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
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_chul = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect
	j = 0
	for i=0 to loop_cnt + 1
'		if code_tab(i) = "" then
'			exit for
'		end if

        Sql = "select * from met_chulgo_goods where (chulgo_date = '"&chul_date&"') and (chulgo_stock = '"&stock_code&"') and (chulgo_seq = '"&chul_seq&"') and (cg_goods_code = '"&code_tab(i)&"')"
		Set Rs_chul=DbConn.Execute(Sql)
		if not Rs_chul.eof then
		       cg_qty = Rs_chul("cg_qty")
		   else
		       cg_qty = 0
	    end if
'		Sql="select * from met_stock_gmaster where stock_code = '" + stock_code + "' and stock_goods_type = '" + goods_type + "' and stock_goods_code ='"+code_tab(i)+"'"
		Sql="select * from met_stock_gmaster where stock_code = '" + stock_code + "' and stock_goods_code ='"+code_tab(i)+"'"
		Set rs=DbConn.Execute(Sql)
		if rs.eof or rs.bof then
			response.write"<script language=javascript>"
			response.write"opener.document.frm.srv_type"&i&".value = '';"
			response.write"opener.document.frm.goods_gubun"&i&".value = '';"
			response.write"opener.document.frm.goods_code"&i&".value = '';"
			response.write"opener.document.frm.goods_name"&i&".value = '';"
			response.write"opener.document.frm.goods_standard"&i&".value = '';"
			response.write"opener.document.frm.goods_grade"&i&".value = '';"
			response.write"opener.document.frm.jqty"&i&".value = '0';"
'			response.write"opener.document.frm.qty"&i&".value = '0';"
            response.write"opener.document.frm.qty"&i&".value = '"&cg_qty&"';"
			response.write"opener.document.frm.del_check"&i&".checked = false;"
			response.write"opener.document.getElementById('pummok_list"&i&"').style.display = 'none';"
			response.write"</script>"
		  else
			response.write"<script language=javascript>"
			response.write"opener.document.frm.srv_type"&i&".value = '"&rs("stock_goods_type")&"';"
		    response.write"opener.document.frm.goods_gubun"&i&".value = '"&rs("stock_goods_gubun")&"';"
		    response.write"opener.document.frm.goods_code"&i&".value = '"&rs("stock_goods_code")&"';"
		    response.write"opener.document.frm.goods_name"&i&".value = '"&rs("stock_goods_name")&"';"
		    response.write"opener.document.frm.goods_standard"&i&".value = '"&rs("stock_goods_standard")&"';"
		    response.write"opener.document.frm.jqty"&i&".value = '"&rs("stock_JJ_qty")&"';"
			response.write"opener.document.frm.qty"&i&".value = '"&cg_qty&"';"
			response.write"opener.document.frm.goods_grade"&i&".value = '"&rs("stock_goods_grade")&"';"
			response.write"opener.document.frm.del_check"&i&".checked = false;"
			response.write"opener.document.getElementById('pummok_list"&i&"').style.display = '';"
			response.write"</script>"
		end if
	next
	response.write"<script language=javascript>"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing				
%>

