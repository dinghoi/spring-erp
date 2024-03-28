<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	u_type = abc("u_type")
	
	old_chulgo_date = abc("old_chulgo_date")
	old_chulgo_stock = abc("old_chulgo_stock")
	old_chulgo_seq = abc("old_chulgo_seq")
	old_chulgo_goods_type = abc("old_chulgo_goods_type")
	old_chulgo_att_file = abc("old_chulgo_att_file")
	old_rele_stock = abc("old_rele_stock")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
	Set Rs_chul = Server.CreateObject("ADODB.Recordset")
	Set Rs_mvin = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	Set Rs_rele = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans

' 재고조정
        Sql = "select * from met_chulgo_goods where (chulgo_date = '"&old_chulgo_date&"') and (chulgo_stock = '"&old_chulgo_stock&"') and (chulgo_seq = '"&old_chulgo_seq&"')"
	    Set Rs_chul=DbConn.Execute(Sql)
		do until Rs_chul.eof
            mod_stock_code = Rs_chul("chulgo_stock")
			mod_goods_type = Rs_chul("cg_goods_type")
			mod_goods_code = Rs_chul("cg_goods_code")
			
			mod_chul_qty = Rs_chul("cg_qty")
' 출고창고 재고정리     
			sql="select * from met_stock_gmaster where stock_code='"&mod_stock_code&"' and stock_goods_code='"&mod_goods_code&"' and stock_goods_type='"&mod_goods_type&"'"
	        set Rs_jago=dbconn.execute(sql)

            if not Rs_jago.eof then
			       go_a_qty = Rs_jago("stock_go_qty")
				   JJ_a_qty = Rs_jago("stock_JJ_qty")
							 
				   go_a_qty = go_a_qty - mod_chul_qty
				   JJ_a_qty = JJ_a_qty + mod_chul_qty
							 
	               sql = "update met_stock_gmaster set stock_go_qty='"&go_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&mod_stock_code&"' and stock_goods_type='"&mod_goods_type&"' and stock_goods_code='"&mod_goods_code&"'"

		          'response.write sql
		
		           dbconn.execute(sql)	  
            end if	 
' 요청(입고)창고 재고정리    
'response.write(old_rele_stock)
			sql="select * from met_stock_gmaster where stock_code='"&old_rele_stock&"' and stock_goods_type='"&mod_goods_type&"' and stock_goods_code='"&mod_goods_code&"'"
	        set Rs_rele=dbconn.execute(sql)

            if not Rs_rele.eof then
				   in_a_qty = Rs_rele("stock_in_qty")
				   JJ_a_qty = Rs_rele("stock_JJ_qty")
							 
				   in_a_qty = in_a_qty - mod_chul_qty
				   JJ_a_qty = JJ_a_qty - mod_chul_qty
							 
	               sql = "update met_stock_gmaster set stock_in_qty='"&in_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&old_rele_stock&"' and stock_goods_type='"&mod_goods_type&"' and stock_goods_code='"&mod_goods_code&"'"

		          'response.write sql
		
		           dbconn.execute(sql)	  
            end if	 

	 
		    Rs_chul.movenext()
	    loop
        Rs_chul.close()
' 입고정리
        sql="select * from met_mv_in where chulgo_date='"&old_chulgo_date&"' and chulgo_stock='"&old_chulgo_stock&"' and chulgo_seq='"&old_chulgo_seq&"'"
	    set Rs_mvin=dbconn.execute(sql)

        if not Rs_mvin.eof then
		       mvin_in_date = Rs_mvin("mvin_in_date")
			   mvin_in_stock = Rs_mvin("mvin_in_stock")
			   mvin_in_seq = Rs_mvin("mvin_in_seq")
		
		       sql = "delete from met_mv_in where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"')"
		       dbconn.execute(sql)
		       sql = "delete from met_mv_in_goods where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"')"
		       dbconn.execute(sql)
		end if
' 출고정리		
		sql = "delete from met_chulgo where (chulgo_date = '"&old_chulgo_date&"') and (chulgo_stock = '"&old_chulgo_stock&"') and (chulgo_seq = '"&old_chulgo_seq&"')"
		dbconn.execute(sql)
		sql = "delete from met_chulgo_goods where (chulgo_date = '"&old_chulgo_date&"') and (chulgo_stock = '"&old_chulgo_stock&"') and (chulgo_seq = '"&old_chulgo_seq&"')"
		dbconn.execute(sql)
		

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
