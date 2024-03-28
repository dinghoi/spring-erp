<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	slip_id = request("slip_id")
	slip_no = request("slip_no")
	slip_seq = request("slip_seq")
	slip_stat = request("slip_stat")

	dbconn.BeginTrans

	if slip_stat = "2" or slip_stat = "3" then
' 기존 대기전표에 수주된 금액 업데이트	
		sql = "select * from sales_slip where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
		set rs=dbconn.execute(sql)
		sales_price = int(rs("sales_price"))	
		sales_cost = int(rs("sales_cost"))		
		sales_cost_vat = int(rs("sales_cost_vat"))
		rs.close()

		sql = "select * from sales_slip where slip_no = '"&slip_no&"' and slip_id = '1'"
		set rs=dbconn.execute(sql)
		order_price = int(rs("order_price")) - sales_price			
		order_cost = int(rs("order_cost")) - sales_cost			
		order_cost_vat = int(rs("order_cost_vat")) - sales_cost_vat			
		if order_price = 0 then
			cal_stat = 1 
		  else
		  	cal_stat = 2
		end if
		rs.close()

		sql = "Update sales_slip set slip_stat ='"&cal_stat&"', order_price ="&order_price&", order_cost ="&order_cost&", order_cost_vat ="&order_cost_vat&"  where slip_no = '"&slip_no&"' and slip_id = '1'"
		dbconn.execute(sql)

		sql = "select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
		Rs.Open Sql, Dbconn, 1
		do until rs.eof
			sql = "select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '1' and goods_seq = '"&rs("goods_seq")&"'"
			set rs_detail=dbconn.execute(sql)
			order_qty = int(rs_detail("order_qty")) - int(rs("qty"))			

			sql = "Update sales_slip_detail set order_qty ="&order_qty&" where slip_no = '"&slip_no&"' and slip_id = '1' and goods_seq = '"&rs("goods_seq")&"'"
			dbconn.execute(sql)

			rs.movenext()
		loop
	end if

	sql = "delete from sales_slip where slip_id ='"&slip_id&"' and slip_no='"&slip_no&"' and slip_seq='"&slip_seq&"'"
	dbconn.execute(sql)
	sql = "delete from sales_slip_detail where slip_id ='"&slip_id&"' and slip_no='"&slip_no&"' and slip_seq='"&slip_seq&"'"
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
	response.write"location.replace('sales_slip_ing_mg.asp');"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
