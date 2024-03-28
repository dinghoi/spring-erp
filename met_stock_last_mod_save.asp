<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	stock_code = request.form("stock_code")
	stock_goods_type = request.form("stock_goods_type")
	stock_goods_code = request.form("stock_goods_code")
	
	stock_last_qty = int(request.form("stock_last_qty"))
	stock_last_amt = int(request.form("stock_last_amt"))
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans

    emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
' 재고 등록				 
        sql="select * from met_stock_gmaster where stock_code='"&stock_code&"' and stock_goods_code='"&stock_goods_code&"' and stock_goods_type='"&stock_goods_type&"'"
	    set Rs_jago=dbconn.execute(sql)
        if not Rs_jago.eof then
				st_in_qty = Rs_jago("stock_in_qty")
				st_in_amt = Rs_jago("stock_in_amt")
				st_go_qty = Rs_jago("stock_go_qty")
				st_go_amt = Rs_jago("stock_go_amt")
							 
				jj_a_qty = stock_last_qty + st_in_qty - st_go_qty
				jj_a_amt = stock_last_amt + st_in_amt - st_go_amt
		
		        sql = "update met_stock_gmaster set stock_last_qty='"&stock_last_qty&"',stock_last_amt='"&stock_last_amt&"',stock_JJ_qty='"&jj_a_qty&"',stock_jj_amt='"&jj_a_amt&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&stock_code&"' and stock_goods_type='"&stock_goods_type&"' and stock_goods_code='"&stock_goods_code&"'"
				 
'		response.write sql
		
        dbconn.execute(sql)	  
		end if
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
