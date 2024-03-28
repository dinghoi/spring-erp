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
	dim oqty_tab(20)
	dim qty_tab(20)
	dim buy_cost(20)
	dim buy_amt(20)
	dim seq_tab(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
	    oqty_tab(i) = 0
		qty_tab(i) = 0
	    buy_cost(i) = 0
	    buy_amt(i) = 0
	    seq_tab(i) = ""
    next
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	u_type = request.form("u_type")

	buy_no = request.form("buy_no")
	buy_seq = request.form("buy_seq")
	buy_date = request.form("buy_date")
	
	buy_goods_type = request.form("buy_goods_type")
	
	order_id = request.form("order_id")
	order_no = request.form("order_no")
	order_seq = request.form("order_seq")
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
	
	stin_in_date = request.form("stin_in_date")
    stin_stock_company = request.form("stin_stock_company")
	stin_stock_code = request.form("stin_stock_code")
    stin_stock_name = request.form("stin_stock_name")
	
	stin_bill_collect = request.form("bill_collect")
	stin_collect_due_date = request.form("collect_due_date")
	
	stin_price = int(request.form("buy_tot_price"))
	stin_cost = int(request.form("buy_tot_cost"))
	stin_cost_vat = int(request.form("buy_tot_cost_vat"))
	
	stin_emp_no = request.form("emp_no")
    stin_emp_name = request.form("emp_name")
	stin_company = request.form("emp_company")
    stin_org_name = request.form("emp_org_name")
    
	order_ing = "4"
	stin_id = "구매입고"
	if order_id = "2" then
		stin_id = "수주전표"
    end if
    if order_id = "1" then
		stin_id = "대기전표"
    end if
	stin_type = "정상"
	
	order_cost = 0
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
			oqty_tab(i) = int(request.form("oqty"&i))
		    qty_tab(i) = int(request.form("qty"&i))
		    buy_cost(i) = int(request.form("buy_cost"&i))
		    buy_amt(i) = int(request.form("buy_tot"&i))
			order_cost = order_cost + buy_amt(i)
		end if
	next
	order_cost_vat = Int(order_cost * (10 / 100))
	order_price = order_cost + order_cost_vat
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
	Set Rs_goods = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	
	dbconn.BeginTrans

	yymmdd = mid(cstr(stin_in_date),3,2) + mid(cstr(stin_in_date),6,2)  + mid(cstr(stin_in_date),9,2)
	
	Sql = "SELECT * FROM met_stock_code where stock_code = '"&stin_stock_code&"'"
    Set Rs_stock = DbConn.Execute(SQL)
    if not Rs_stock.eof then
       	   stock_level = Rs_stock("stock_level")
		   stock_name = Rs_stock("stock_name")
		   stock_company = Rs_stock("stock_company")
		   stock_bonbu = Rs_stock("stock_bonbu")
		   stock_saupbu = Rs_stock("stock_saupbu")
		   stock_team = Rs_stock("stock_team")
        else
		   stock_level = ""
		   stock_name = ""
		   stock_company = ""
		   stock_bonbu = ""
		   stock_saupbu = ""
		   stock_team = ""
    end if
    Rs_stock.close()
	
'구매 update
    sql = "Update met_buy set buy_ing='"&order_ing&"',mod_date=now(),mod_user='"&user_name&"' where buy_no = '"&buy_no&"' and buy_seq = '"&buy_seq&"' and buy_date = '"&buy_date&"'"
		dbconn.execute(sql)

'구매 품목 update
    for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		  	     sql = "Update met_buy_goods set bg_ing='"&order_ing&"',mod_date=now(),mod_user='"&user_name&"' where bg_no = '"&buy_no&"' and buy_seq = '"&buy_seq&"' and bg_date = '"&buy_date&"' and bg_goods_code = '"&code_tab(i)&"'"
			     dbconn.execute(sql)
		end if
	next

'발주 update
    sql = "Update met_order set order_ing='"&order_ing&"',mod_date=now(),mod_user='"&user_name&"' where order_no = '"&order_no&"' and order_seq = '"&order_seq&"' and order_date = '"&order_date&"'"
		dbconn.execute(sql)

'발주 품목 update
    for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		  	     sql = "Update met_order_goods set og_ing='"&order_ing&"',mod_date=now(),mod_user='"&user_name&"' where og_order_no = '"&order_no&"' and og_order_seq = '"&order_seq&"' and og_order_date = '"&order_date&"' and og_seq = '"&seq_tab(i)&"' and og_goods_code = '"&code_tab(i)&"'"
			     dbconn.execute(sql)
		end if
	next
	
'입고등록	
	
	sql="insert into met_stin (stin_in_date,stin_order_no,stin_order_seq,stin_order_date,stin_buy_no,stin_buy_seq,stin_buy_date,stin_goods_type,stin_trade_no,stin_trade_name,stin_trade_person,stin_trade_email,stin_stock_company,stin_stock_code,stin_stock_name,stin_id,stin_type,stin_bill_collect,stin_collect_due_date,stin_price,stin_cost,stin_cost_vat,stin_company,stin_org_name,stin_emp_no,stin_emp_name,reg_date,reg_user) values ('"&stin_in_date&"','"&order_no&"','"&order_seq&"','"&order_date&"','"&buy_no&"','"&buy_seq&"','"&buy_date&"','"&buy_goods_type&"','"&order_trade_no&"','"&order_trade_name&"','"&order_trade_person&"','"&order_trade_email&"','"&stin_stock_company&"','"&stin_stock_code&"','"&stin_stock_name&"','"&stin_id&"','"&stin_type&"','"&stin_bill_collect&"','"&stin_collect_due_date&"','"&stin_price&"','"&stin_cost&"','"&stin_cost_vat&"','"&stin_company&"','"&stin_org_name&"','"&stin_emp_no&"','"&stin_emp_name&"',now(),'"&user_name&"')"

	dbconn.execute(sql)
	
'입고 품목 등록	
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			     goods_seq = right(("00" + cstr(i)),2)
		  	     sql="insert into met_stin_goods (stin_date,stin_order_no,stin_order_seq,stin_goods_code,stin_goods_seq,stin_goods_type,stin_goods_gubun,stin_goods_name,stin_standard,stin_unit_cost,stin_qty,stin_amt,stin_id,stin_type,stin_stock_company,stin_stock_code,stin_stock_name,reg_date,reg_user) values ('"&stin_in_date&"','"&order_no&"','"&order_seq&"','"&code_tab(i)&"','"&goods_seq&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&goods_standard(i)&"','"&buy_cost(i)&"','"&qty_tab(i)&"','"&buy_amt(i)&"','"&stin_id&"','"&stin_type&"','"&stin_stock_company&"','"&stin_stock_code&"','"&stin_stock_name&"',now(),'"&user_name&"')"
			     dbconn.execute(sql)
				 
' 재고 등록				 
                 Sql = "SELECT * FROM met_goods_code where goods_type = '"&goods_type(i)&"' and goods_code = '"&code_tab(i)&"'"
                 Set Rs_goods = DbConn.Execute(SQL)
                 if not Rs_goods.eof then
       	                goods_grade = Rs_goods("goods_grade")
                    else
		                goods_grade = ""
                 end if
                 Rs_goods.close()
				 
				 stock_goods_level1 = mid(cstr(code_tab(i)),1,3)
                 stock_goods_level2 = mid(cstr(code_tab(i)),4,4)
                 stock_goods_seq = mid(cstr(code_tab(i)),8,3) 
				 
				 sql="select * from met_stock_gmaster where stock_code='"&stin_stock_code&"' and stock_goods_code='"&code_tab(i)&"' and stock_goods_type='"&goods_type(i)&"'"
	             set Rs_jago=dbconn.execute(sql)

                      if Rs_jago.eof then
                             sql = "insert into met_stock_gmaster(stock_code,stock_goods_type,stock_goods_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_goods_level1,stock_goods_level2,stock_goods_seq,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_in_qty,stock_in_amt,stock_JJ_qty,stock_jj_amt,reg_date,reg_user) values "
	                         sql = sql +	" ('"&stin_stock_code&"','"&goods_type(i)&"','"&code_tab(i)&"','"&stock_level&"','"&stock_name&"','"&stock_company&"','"&stock_bonbu&"','"&stock_saupbu&"','"&stock_team&"','"&stock_goods_level1&"','"&stock_goods_level2&"','"&stock_goods_seq&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&goods_standard(i)&"','"&goods_grade&"','"&qty_tab(i)&"','"&buy_amt(i)&"','"&qty_tab(i)&"','"&buy_amt(i)&"',now(),'"&user_name&"')"

		'response.write(sql)
		                     dbconn.execute(sql)	 
	                     else
						     in_a_qty = Rs_jago("stock_in_qty")
							 in_a_amt = Rs_jago("stock_in_amt")
							 JJ_a_qty = Rs_jago("stock_JJ_qty")
							 jj_a_amt = Rs_jago("stock_jj_amt")
							 
							 in_a_qty = in_a_qty + qty_tab(i)
							 in_a_amt = in_a_amt + buy_amt(i)
							 JJ_a_qty = JJ_a_qty + qty_tab(i)
							 jj_a_amt = jj_a_amt + buy_amt(i)
							 
	                         sql = "update met_stock_gmaster set stock_in_qty='"&in_a_qty&"',stock_in_amt='"&in_a_amt&"',stock_JJ_qty='"&JJ_a_qty&"',stock_jj_amt='"&jj_a_amt&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&stin_stock_code&"' and stock_goods_type='"&goods_type(i)&"' and stock_goods_code='"&code_tab(i)&"'"

		'response.write sql
		
		                     dbconn.execute(sql)	  
                       end if
		end if
	next

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"location.replace('meterials_stock_out_mg.asp');"
	response.write"self.opener.location.reload();"	
	response.write"window.close();"	
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
