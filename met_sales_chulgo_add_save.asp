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
	dim goods_grade(20)
	dim qty_tab(20)
	dim seq_tab(20)
	dim c_qty_tab(20)
	dim chul_qty_tab(20)
	dim chul_qty_hap(20)
	dim chulgo_ing_tab(20)
	dim c_chk_tab(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
		goods_grade(i) = ""
		seq_tab(i) = ""
		qty_tab(i) = 0
	    c_qty_tab(i) = 0
	    chul_qty_tab(i) = 0
		chul_qty_hap(i) = 0
		chulgo_ing_tab(i) = ""
		c_chk_tab(i) = ""
    next
	
	curr_date = mid(cstr(now()),1,10)
	u_type = request.form("u_type")
	
    slip_id = request.form("slip_id")
	slip_no = request.form("slip_no")
	slip_seq = request.form("slip_seq")
	sales_date = request.form("sales_date")
	slip_id_view = request.form("slip_id_view")
	chulgo_goods_type = ""
'	chulgo_goods_type = rele_goods_type
	
	chulgo_id = slip_id_view
	
	service_no = ""
	chulgo_trade_name = request.form("trade_name")
'   order_trade_no = request.form("trade_no")
'	order_trade_no = replace(request.form("trade_no"),"-","")
'	order_trade_person = request.form("trade_person")
'	order_trade_email = request.form("trade_email")
	chulgo_trade_dept = ""
	slip_memo = request.form("slip_memo")
	chulgo_memo = slip_memo
	
'출고내용	
	chulgo_date = request.form("chulgo_date")
    chulgo_stock = request.form("chulgo_stock")
    chulgo_stock_company = request.form("chulgo_stock_company")
    chulgo_stock_name = request.form("chulgo_stock_name")
    chulgo_emp_no = request.form("chulgo_emp_no")
    chulgo_emp_name = request.form("chulgo_emp_name")
	chulgo_company = request.form("chulgo_company")
	chulgo_bonbu = request.form("chulgo_bonbu")
	chulgo_saupbu = request.form("chulgo_saupbu")
	chulgo_team = request.form("chulgo_team")
	chulgo_org_name = request.form("chulgo_org_name")
	
	chulgo_ing = "출고완료"
	
	for i = 1 to 20	
		code_tab(i) = request.form("goods_code"&i)
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		    seq_tab(i) = request.form("bg_seq"&i)
			c_chk_tab(i) = request.form("c_chk"&i)
			goods_type(i) = request.form("srv_type"&i)
		    goods_gubun(i) = request.form("goods_gubun"&i)
		    goods_name(i) = request.form("goods_name"&i)
		    goods_standard(i) = request.form("goods_standard"&i)
			goods_grade(i) = request.form("goods_grade"&i)
		    qty_tab(i) = int(request.form("qty"&i))
			c_qty_tab(i) = int(request.form("c_qty"&i))
		    chul_qty_tab(i) = int(request.form("chul_qty"&i))
			
			chul_qty_hap(i) = c_qty_tab(i) + chul_qty_tab(i)
		    if qty_tab(i) > chul_qty_hap(i) then 
		           chulgo_ing_tab(i) = "부분출고"
			 	   chulgo_type = "부분출고"
		       else 
		           chulgo_ing_tab(i) = "출고완료"
		    end if
		end if
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	dbconn.BeginTrans

	yymm = mid(cstr(now()),3,2) + mid(cstr(now()),6,2) 
	
	sql="select max(chulgo_seq) as max_seq from met_chulgo where chulgo_date = '"&chulgo_date&"' and chulgo_stock = '"&chulgo_stock&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_seq = "01"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		code_seq = right(max_seq,2)
	end if
    rs_max.close()
	
if u_type = "U" then
	   code_last = chulgo_seq
   else
       code_last = code_seq
end if
	
chulgo_seq = code_last

'출고등록	
	
	sql="insert into met_chulgo (chulgo_date,chulgo_stock,chulgo_seq,chulgo_goods_type,chulgo_id,service_no,chulgo_trade_name,chulgo_trade_dept,chulgo_stock_company,chulgo_stock_name,chulgo_emp_no,chulgo_emp_name,chulgo_company,chulgo_bonbu,chulgo_saupbu,chulgo_team,chulgo_org_name,chulgo_memo,rele_no,rele_seq,rele_date,reg_date,reg_user) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&chulgo_goods_type&"','"&chulgo_id&"','"&service_no&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&chulgo_emp_no&"','"&chulgo_emp_name&"','"&chulgo_company&"','"&chulgo_bonbu&"','"&chulgo_saupbu&"','"&chulgo_team&"','"&chulgo_org_name&"','"&chulgo_memo&"','"&slip_no&"','"&slip_seq&"','"&sales_date&"',now(),'"&user_name&"')"

	dbconn.execute(sql)
	
'출고 품목 등록	
	j = 0
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		    if c_chk_tab(i) <> "1" and chul_qty_tab(i) <> 0 then
			     j = j + 1
				 cg_goods_seq = right(("00" + cstr(j)),2)
		  	     sql="insert into met_chulgo_goods (chulgo_date,chulgo_stock,chulgo_seq,cg_goods_seq,cg_goods_code,cg_goods_type,cg_goods_gubun,cg_standard,cg_goods_name,cg_goods_grade,cg_qty,cg_type,rl_trade_name,reg_date,reg_user) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&cg_goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"','"&chul_qty_tab(i)&"','"&chulgo_id&"','"&chulgo_trade_name&"',now(),'"&user_name&"')"
				 
			     dbconn.execute(sql)
				 
' 재고 등록				 
			 
				 sql="select * from met_stock_gmaster where stock_code='"&chulgo_stock&"' and stock_goods_code='"&code_tab(i)&"' and stock_goods_type='"&goods_type(i)&"'"
	             set Rs_jago=dbconn.execute(sql)

                      if not Rs_jago.eof then
						     go_a_qty = Rs_jago("stock_go_qty")
							 JJ_a_qty = Rs_jago("stock_JJ_qty")
							 
							 go_a_qty = go_a_qty + chul_qty_tab(i)
							 JJ_a_qty = JJ_a_qty - chul_qty_tab(i)
							 
	                         sql = "update met_stock_gmaster set stock_go_qty='"&go_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&chulgo_stock&"' and stock_goods_type='"&goods_type(i)&"' and stock_goods_code='"&code_tab(i)&"'"
				 
		'response.write sql
		
		                     dbconn.execute(sql)	  
                       end if				 
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
