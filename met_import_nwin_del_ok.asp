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
	
	old_stin_in_date = abc("old_stin_in_date")
	old_stin_order_no = abc("old_stin_order_no")
	old_stin_order_seq = abc("old_stin_order_seq")
	old_stin_goods_type = abc("old_stin_goods_type")
	old_stin_att_file = abc("old_stin_att_file")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_stin = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans

' 재고조정
        Sql="select * from met_stin_goods where (stin_date = '"&old_stin_in_date&"') and (stin_order_no = '"&old_stin_order_no&"') and (stin_order_seq = '"&old_stin_order_seq&"')"
	    Set Rs_stin=DbConn.Execute(Sql)
		do until Rs_stin.eof
            mod_stock_code = Rs_stin("stin_stock_code")
			mod_goods_type = Rs_stin("stin_goods_type")
			mod_goods_code = Rs_stin("stin_goods_code")
			mod_stock_company = Rs_stin("stin_stock_company")
			
			mod_stin_qty = Rs_stin("stin_qty")
			mod_stin_amt = Rs_stin("stin_amt")
     
			sql="select * from met_stock_gmaster where stock_code='"&mod_stock_code&"' and stock_goods_code='"&mod_goods_code&"' and stock_goods_type='"&mod_goods_type&"'"
	        set Rs_jago=dbconn.execute(sql)

            if not Rs_jago.eof then
			       in_a_qty = Rs_jago("stock_in_qty")
				   in_a_amt = Rs_jago("stock_in_amt")
				   jj_a_qty = Rs_jago("stock_JJ_qty")
				   jj_a_amt = Rs_jago("stock_jj_amt")
							 
				   in_a_qty = in_a_qty - mod_stin_qty
				   in_a_amt = in_a_amt - mod_stin_amt
				   jj_a_qty = jj_a_qty - mod_stin_qty
				   jj_a_amt = jj_a_amt - mod_stin_amt
							 
	               sql = "update met_stock_gmaster set stock_in_qty='"&in_a_qty&"',stock_in_amt='"&in_a_amt&"',stock_JJ_qty='"&JJ_a_qty&"',stock_jj_amt='"&jj_a_amt&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&mod_stock_code&"' and stock_goods_type='"&mod_goods_type&"' and stock_goods_code='"&mod_goods_code&"'"

		          'response.write sql
		
		           dbconn.execute(sql)	  
            end if	 
	 
		    Rs_stin.movenext()
	    loop
        Rs_stin.close()

		sql = "delete from met_stin where stin_in_date ='"&old_stin_in_date&"' and stin_order_no='"&old_stin_order_no&"' and stin_order_seq='"&old_stin_order_seq&"'"
		dbconn.execute(sql)
		sql = "delete from met_stin_goods where stin_date ='"&old_stin_in_date&"' and stin_order_no='"&old_stin_order_no&"' and stin_order_seq='"&old_stin_order_seq&"'"
		dbconn.execute(sql)
' serial no삭제		
		sql = "delete from met_goods_serial where in_date ='"&old_stin_in_date&"' and in_order_no='"&old_stin_order_no&"' and in_order_seq='"&old_stin_order_seq&"'"
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
