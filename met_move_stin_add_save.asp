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
	dim seq_tab(20)
	dim qty_tab(20)
	dim c_qty_tab(20)
	dim mvin_qty_tab(20)
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
	    mvin_qty_tab(i) = 0
		c_chk_tab(i) = ""
    next
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	u_type = request.form("u_type")
'출고의뢰 내용
	rele_date = request.form("rele_date")
	rele_stock = request.form("rele_stock")
	rele_seq = request.form("rele_seq")
	
	rele_stock_company = request.form("rele_stock_company")
    rele_stock_name = request.form("rele_stock_name")
	
'출고내용	
	chulgo_date = request.form("chulgo_date")
    chulgo_stock = request.form("chulgo_stock")
	chulgo_seq = request.form("chulgo_seq")
	chulgo_type = request.form("chulgo_type")
	chulgo_goods_type = request.form("chulgo_goods_type")
	chulgo_memo = request.form("chulgo_memo")
	
'	if chulgo_type = "출고완료" then
	     chulgo_type = "입고완료"
'	end if
	
'입고내용	
	mvin_in_stock = rele_stock
	mvin_stock_company = rele_stock_company
	mvin_stock_name = rele_stock_name
	mvin_goods_type = chulgo_goods_type
	mvin_in_date = request.form("mvin_in_date")
    mvin_emp_no = request.form("mvin_emp_no")
	mvin_emp_name = request.form("mvin_emp_name")
	mvin_company = request.form("mvin_company")
	mvin_bonbu = request.form("mvin_bonbu")
	mvin_saupbu = request.form("mvin_saupbu")
	mvin_team = request.form("mvin_team")
	mvin_org_name = request.form("mvin_org_name")
	mvin_id = "창고이동"
	
	for i = 1 to 20	
		code_tab(i) = request.form("goods_code"&i)
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		    c_chk_tab(i) = request.form("c_chk"&i)
			seq_tab(i) = request.form("goods_seq"&i)
			goods_type(i) = request.form("srv_type"&i)
		    goods_gubun(i) = request.form("goods_gubun"&i)
		    goods_name(i) = request.form("goods_name"&i)
		    goods_standard(i) = request.form("goods_standard"&i)
			goods_grade(i) = request.form("goods_grade"&i)
		    qty_tab(i) = int(request.form("qty"&i))
			c_qty_tab(i) = int(request.form("c_qty"&i))
		    mvin_qty_tab(i) = int(request.form("mvin_qty"&i))
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
	
	Sql = "SELECT * FROM met_stock_code where stock_code = '"&mvin_in_stock&"'"
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
  
	sql="select max(mvin_in_seq) as max_seq from met_mv_in where mvin_in_date = '"&mvin_in_date&"' and mvin_in_stock = '"&mvin_in_stock&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_seq = "01"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		code_seq = right(max_seq,2)
	end if
    rs_max.close()
	
if u_type = "U" then
	   code_last = mvin_in_seq
   else
       code_last = code_seq
end if
	
mvin_in_seq = code_last

'출고의뢰 update
    sql = "Update met_mv_reg set chulgo_ing='"&chulgo_type&"',in_stock_date='"&mvin_in_date&"',mod_date=now(),mod_user='"&user_name&"' where rele_date = '"&rele_date&"' and rele_stock = '"&rele_stock&"' and rele_seq = '"&rele_seq&"'"
		dbconn.execute(sql)

'출고의뢰 품목 update
'    for i = 1 to 20
'		if code_tab(i) = "" or isnull(code_tab(i)) then
'			exit for
'		  else
'		    if c_chk_tab(i) <> "1" then
'		  	     sql = "Update met_mv_reg_goods set in_stock_date='"&mvin_in_date&"',mod_date=now(),mod_user='"&user_name&"' where rele_date = '"&rele_date&"' and rele_stock = '"&rele_stock&"' and rele_stock_seq = '"&rele_seq&"' and rele_goods = '"&code_tab(i)&"'"
'			     dbconn.execute(sql)
'			end if
'		end if
'	next
	
'출고 입고일update	
	
	sql = "Update met_mv_go set chulgo_type='"&chulgo_type&"',in_stock_date='"&mvin_in_date&"',mod_date=now(),mod_user='"&user_name&"' where chulgo_date = '"&chulgo_date&"' and chulgo_stock = '"&chulgo_stock&"' and chulgo_seq = '"&chulgo_seq&"'"

	dbconn.execute(sql)
	
'출고 품목 입고일update
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
'		    if c_chk_tab(i) <> "1" then
		  	     sql = "Update met_mv_go_goods set cg_type='"&chulgo_type&"',mod_date=now(),mod_user='"&user_name&"' where chulgo_date = '"&chulgo_date&"' and chulgo_stock = '"&chulgo_stock&"' and chulgo_seq = '"&chulgo_seq&"' and cg_goods_code = '"&code_tab(i)&"' and cg_goods_seq = '"&seq_tab(i)&"'"
			     dbconn.execute(sql)
'			end if
		end if
	next


'입고등록	
	
	sql="insert into met_mv_in (mvin_in_date,mvin_in_stock,mvin_in_seq,mvin_id,mvin_goods_type,mvin_stock_company,mvin_stock_name,mvin_emp_no,mvin_emp_name,mvin_company,mvin_bonbu,mvin_saupbu,mvin_team,mvin_org_name,rele_date,rele_stock,rele_seq,chulgo_date,chulgo_stock,chulgo_seq,reg_date,reg_user) values ('"&mvin_in_date&"','"&mvin_in_stock&"','"&mvin_in_seq&"','"&mvin_id&"','"&mvin_goods_type&"','"&mvin_stock_company&"','"&mvin_stock_name&"','"&mvin_emp_no&"','"&mvin_emp_name&"','"&mvin_company&"','"&mvin_bonbu&"','"&mvin_saupbu&"','"&mvin_team&"','"&mvin_org_name&"','"&rele_date&"','"&rele_stock&"','"&rele_seq&"','"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"',now(),'"&user_name&"')"

	dbconn.execute(sql)

	
'입고 품목 등록	
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
'		    if c_chk_tab(i) <> "1" and chul_qty_tab(i) <> 0 then
			     goods_seq = right(("00" + cstr(i)),2)
		  	     sql="insert into met_mv_in_goods (mvin_in_date,mvin_in_stock,mvin_in_seq,in_goods_seq,in_goods_code,in_goods_type,in_goods_gubun,in_standard,in_goods_name,in_goods_grade,in_qty,reg_date,reg_user) values ('"&mvin_in_date&"','"&mvin_in_stock&"','"&mvin_in_seq&"','"&goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"','"&mvin_qty_tab(i)&"',now(),'"&user_name&"')"
			     dbconn.execute(sql)
'			end if

' 재고 등록				 
                 
				stock_goods_level1 = mid(cstr(code_tab(i)),1,3)
                stock_goods_level2 = mid(cstr(code_tab(i)),4,4)
                stock_goods_seq = mid(cstr(code_tab(i)),8,3) 
				 
				 sql="select * from met_stock_gmaster where stock_code='"&mvin_in_stock&"' and stock_goods_type='"&mvin_goods_type&"' and stock_goods_code='"&code_tab(i)&"'"
	             set Rs_jago=dbconn.execute(sql)

                      if Rs_jago.eof then
                             sql = "insert into met_stock_gmaster(stock_code,stock_goods_type,stock_goods_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_goods_level1,stock_goods_level2,stock_goods_seq,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_in_qty,stock_in_amt,stock_JJ_qty,stock_jj_amt,reg_date,reg_user) values "
	                         sql = sql +	" ('"&mvin_in_stock&"','"&mvin_goods_type&"','"&code_tab(i)&"','"&stock_level&"','"&stock_name&"','"&stock_company&"','"&stock_bonbu&"','"&stock_saupbu&"','"&stock_team&"','"&stock_goods_level1&"','"&stock_goods_level2&"','"&stock_goods_seq&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&goods_standard(i)&"','"&goods_grade(i)&"','"&mvin_qty_tab(i)&"',0,'"&mvin_qty_tab(i)&"',0,now(),'"&user_name&"')"

		'response.write(sql)
		                     dbconn.execute(sql)	 
	                     else
						     in_a_qty = Rs_jago("stock_in_qty")
							 JJ_a_qty = Rs_jago("stock_JJ_qty")
							 
							 in_a_qty = in_a_qty + mvin_qty_tab(i)
							 JJ_a_qty = JJ_a_qty + mvin_qty_tab(i)
							 
	                         sql = "update met_stock_gmaster set stock_in_qty='"&in_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&mvin_in_stock&"' and stock_goods_type='"&goods_type(i)&"' and stock_goods_code='"&code_tab(i)&"'"

		'response.write sql
		
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
