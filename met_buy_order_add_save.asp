<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	dim code_tab(20)
	dim goods_name(20)
	dim goods_type(20)
	dim goods_gubun(20)
	dim goods_standard(20)
	dim bqty_tab(20)
	dim qty_tab(20)
	dim buy_cost(20)
	dim buy_amt(20)
	dim seq_tab(20)
	dim oqty_tab(20)
	
	dim ing_type_tab(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
	    bqty_tab(i) = 0
		qty_tab(i) = 0
	    buy_cost(i) = 0
	    buy_amt(i) = 0
	    seq_tab(i) = ""
		oqty_tab(i) = 0
		ing_type_tab(i) = ""
    next
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	u_type = request.form("u_type")
	
	mok_cnt = request.form("mok_cnt")

	order_id = "0"
	buy_no = request.form("buy_no")
	buy_seq = request.form("buy_seq")
	buy_date = request.form("buy_date")
	order_goods_type = request.form("buy_goods_type")
	
	order_date = request.form("order_date")
    order_company = request.form("order_company")
	order_bonbu = request.form("order_bonbu")
	order_saupbu = request.form("order_saupbu")
	order_team = request.form("order_team")
	order_org_code = request.form("order_org_code")
	order_org_name = request.form("order_org_name")
	order_emp_no = request.form("order_emp_no")
    order_emp_name = request.form("order_emp_name")
	order_trade_name = request.form("trade_name")
'    order_trade_no = request.form("trade_no")
	order_trade_no = replace(request.form("trade_no"),"-","")
	order_trade_person = request.form("trade_person")
	order_trade_email = request.form("trade_email")
	order_bill_collect = request.form("bill_collect")
	order_collect_due_date = request.form("collect_due_date")
	order_in_date = request.form("order_in_date")
    order_stock_company = request.form("order_stock_company")
	order_stock_code = request.form("order_stock_code")
    order_stock_name = request.form("order_stock_name")
	order_out_method = ""
	order_out_request_date = ""
	if order_collect_due_date = "" or isnull(order_collect_due_date) then
		order_collect_due_date = "0000-00-00"
	end if
	if order_out_request_date = "" or isnull(order_out_request_date) then
		order_out_request_date = "0000-00-00"
	end if
	order_memo = request.form("order_memo")
	order_memo = Replace(order_memo,"'","&quot;")
	
	order_price = int(request.form("buy_tot_price"))
	order_cost = int(request.form("buy_tot_cost"))
	order_cost_vat = int(request.form("buy_tot_cost_vat"))
    
	order_cost = 0
    order_ing = "2" 
    
	for i = 1 to 20	
		code_tab(i) = request.form("goods_code"&i) 
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			goods_type(i) = request.form("srv_type"&i)
			seq_tab(i) = request.form("bg_seq"&i)
		    goods_gubun(i) = request.form("goods_gubun"&i)
		    goods_name(i) = request.form("goods_name"&i)
		    goods_standard(i) = request.form("goods_standard"&i)
			bqty_tab(i) = int(request.form("bqty"&i))       '구매품의수량
'			b_qty(i) = int(request.form("b_qty"&i))       '구매품의수량 - 기발주수량 = 남은구매의뢰수량
		    qty_tab(i) = int(request.form("qty"&i))         '발주수량
		    buy_cost(i) = int(request.form("buy_cost"&i))
		    'buy_amt(i) = int(request.form("buy_tot"&i))
			buy_amt(i) = qty_tab(i) * buy_cost(i)
			oqty_tab(i) = int(request.form("oqty"&i))       '기발주수량
			order_cost = order_cost + buy_amt(i)
			
			gi_qty = oqty_tab(i) + qty_tab(i)
			if gi_qty = 0 then
			      ing_type_tab(i) = "0"
				  order_ing = "1" 
			   else
			      if bqty_tab(i) > gi_qty then
			             ing_type_tab(i) = "1"
				         order_ing = "1" 
			         else	  
				         ing_type_tab(i) = "2"
				  end if
			end if			 
		end if
	next
	order_cost_vat = Int(order_cost * (10 / 100))
	order_price = order_cost + order_cost_vat
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_buy = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	dbconn.BeginTrans

	yymmdd = mid(cstr(order_date),3,2) + mid(cstr(order_date),6,2)  + mid(cstr(order_date),9,2)

if order_cost = 0 then  '발주내역이 없으면 저장 안함
      end_msg = "발주내역이 없습니다...."
   else

'구매 update
    sql = "Update met_buy set buy_ing='"&order_ing&"',mod_date=now(),mod_user='"&user_name&"' where buy_no = '"&buy_no&"' and buy_seq = '"&buy_seq&"' and buy_date = '"&buy_date&"'"
		dbconn.execute(sql)

'구매 품목 update
    for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			  exit for
		  else
		  	  sql = "select * from met_buy_goods where bg_no = '"&buy_no&"' and buy_seq = '"&buy_seq&"' and bg_date = '"&buy_date&"' and bg_seq = '"&seq_tab(i)&"' and bg_goods_code = '"&code_tab(i)&"'"
			  set Rs_buy=dbconn.execute(sql)  
			  if  not Rs_buy.eof then
			      bg_order_qty = Rs_buy("bg_order_qty") + qty_tab(i)
				 
				  sql = "Update met_buy_goods set bg_ing='"&ing_type_tab(i)&"',bg_order_qty='"&bg_order_qty&"',mod_date=now(),mod_user='"&user_name&"' where bg_no = '"&buy_no&"' and buy_seq = '"&buy_seq&"' and bg_date = '"&buy_date&"' and bg_seq = '"&seq_tab(i)&"' and bg_goods_code = '"&code_tab(i)&"'"
			      dbconn.execute(sql)
			  end if
		end if
	next
	
'발주등록	

    order_seq = "00"
	sql="select max(order_no) as max_no from met_order where order_date = '"&order_date&"'"
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_no"))  then
	          order_no = yymmdd + "001"
	    else
	          max_seq = "00" + cstr((int(right(rs("max_no"),3)) + 1))
	          order_no = yymmdd + cstr(right(max_seq,3))
	end if
	
	sql="insert into met_order (order_id,order_no,order_seq,order_date,order_buy_no,order_buy_seq,order_buy_date,order_goods_type,order_company,order_bonbu,order_saupbu,order_team,order_org_code,order_org_name,order_emp_no,order_emp_name,order_trade_no,order_trade_name,order_trade_person,order_trade_email,order_in_date,order_stock_company,order_stock_code,order_stock_name,order_bill_collect,order_collect_due_date,order_out_method,order_out_request_date,order_price,order_cost,order_cost_vat,order_ing,order_memo,reg_date,reg_user) values ('"&order_id&"','"&order_no&"','"&order_seq&"','"&order_date&"','"&buy_no&"','"&buy_seq&"','"&buy_date&"','"&order_goods_type&"','"&order_company&"','"&order_bonbu&"','"&order_saupbu&"','"&order_team&"','"&order_org_code&"','"&order_org_name&"','"&order_emp_no&"','"&order_emp_name&"','"&order_trade_no&"','"&order_trade_name&"','"&order_trade_person&"','"&order_trade_email&"','"&order_in_date&"','"&order_stock_company&"','"&order_stock_code&"','"&order_stock_name&"','"&order_bill_collect&"','"&order_collect_due_date&"','"&order_out_method&"','"&order_out_request_date&"','"&order_price&"','"&order_cost&"','"&order_cost_vat&"','"&order_ing&"','"&order_memo&"',now(),'"&user_name&"')"

	dbconn.execute(sql)
	
'발주 품목 등록	
	j = 0
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			 if qty_tab(i) <> 0 then
			     j = j + 1
				 goods_seq = right(("00" + cstr(j)),2)
		  	     sql="insert into met_order_goods (og_order_id,og_order_no,og_order_seq,og_order_date,og_seq,og_goods_code,og_goods_type,og_goods_gubun,og_standard,og_goods_name,og_bg_qty,og_qty,og_unit_cost,og_amt,og_ing,reg_date,reg_user) values ('"&order_id&"','"&order_no&"','"&order_seq&"','"&order_date&"','"&goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&bqty_tab(i)&"','"&qty_tab(i)&"','"&buy_cost(i)&"','"&buy_amt(i)&"','"&ing_type_tab(i)&"',now(),'"&user_name&"')"
			     dbconn.execute(sql)
			  end if
		end if
	next

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if
	
end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"location.replace('meterials_buy_order_mg.asp');"
	response.write"self.opener.location.reload();"	
	response.write"window.close();"	
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
