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
	dim qty_tab(20)
	dim goods_grade(20)
	dim goods_return(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
		goods_grade(i) = ""
		goods_return(i) = ""
		qty_tab(i) = 0
    next
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	
	u_type = request.form("u_type")

	chulgo_date = request.form("chulgo_date")
    chulgo_stock = request.form("chulgo_stock")
	chulgo_id = "고객출고"
	chulgo_goods_type = request.form("chulgo_goods_type")
	service_no = request.form("service_no")
	chulgo_trade_name = request.form("chulgo_trade_name")
	chulgo_trade_dept = request.form("chulgo_trade_dept")
	
    chulgo_stock_name = request.form("chulgo_stock_name")
	chulgo_stock_company = request.form("chulgo_stock_company")
	
	chulgo_emp_no = request.form("chulgo_emp_no")
    chulgo_emp_name = request.form("chulgo_emp_name")
	chulgo_company = request.form("chulgo_company")
    chulgo_bonbu = request.form("chulgo_bonbu")
    chulgo_saupbu = request.form("chulgo_saupbu")
    chulgo_team = request.form("chulgo_team")
    chulgo_org_name = request.form("chulgo_org_name")
	
	chulgo_ing = "출고완료"
	chulgo_type = "출고완료"
	
	chulgo_memo = request.form("chulgo_memo")
	chulgo_memo = Replace(chulgo_memo,"'","&quot;")
	
	rele_stock = ""
	rele_stock_company = ""
    rele_stock_name = ""
	
    rele_emp_no = ""
    rele_emp_name = ""
	
    rele_company = ""
    rele_bonbu = ""
    rele_saupbu = ""
    rele_team = ""
    rele_org_name = ""
	
	mvin_in_date = request.form("mvin_in_date")
	mvin_in_stock = request.form("mvin_in_stock")
	mvin_in_seq = request.form("mvin_in_seq")
	in_goods_seq = request.form("in_goods_seq")
	in_goods_code = request.form("in_goods_code")
	
	for i = 1 to 20	
'		goods_type(i) = request.form("srv_type"&i)
		goods_type(i) = chulgo_goods_type
		goods_gubun(i) = request.form("goods_gubun"&i)
		code_tab(i) = request.form("goods_code"&i)
		goods_name(i) = request.form("goods_name"&i)
		goods_standard(i) = request.form("goods_standard"&i)
		qty_tab(i) = int(request.form("qty"&i))
		goods_grade(i) = request.form("goods_grade"&i)
		goods_return(i) = request.form("goods_return"&i)
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	Set Rs_mvin = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	dbconn.BeginTrans

	yymmdd = mid(cstr(chulgo_date),3,2) + mid(cstr(chulgo_date),6,2)  + mid(cstr(chulgo_date),9,2)

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
rele_no = yymmdd + chulgo_stock
rele_seq = chulgo_seq

'입고품목 출고수량 update
    j = 1
	sql = "select * from met_mv_in_goods where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"') and (in_goods_seq = '"&in_goods_seq&"') and (in_goods_code = '"&in_goods_code&"')"
	
    set Rs_mvin=dbconn.execute(sql)
    if not Rs_mvin.eof then
 	       cg_qty = Rs_mvin("in_cg_qty")
		   cg_qty = cg_qty + qty_tab(j)
		   
           sql = "Update met_mv_in_goods set chulgo_no='"&rele_no&"',chulgo_seq='"&rele_seq&"',in_cg_qty='"&cg_qty&"' where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"') and (in_goods_seq = '"&in_goods_seq&"') and (in_goods_code = '"&in_goods_code&"')"
	
	       dbconn.execute(sql)
    end if

'출고등록	

	sql="insert into met_chulgo (chulgo_date,chulgo_stock,chulgo_seq,chulgo_goods_type,chulgo_id,service_no,chulgo_trade_name,chulgo_trade_dept,chulgo_type,chulgo_stock_company,chulgo_stock_name,chulgo_emp_no,chulgo_emp_name,chulgo_company,chulgo_bonbu,chulgo_saupbu,chulgo_team,chulgo_org_name,chulgo_memo,rele_no,rele_seq,rele_stock,rele_stock_company,rele_stock_name,rele_company,rele_bonbu,rele_saupbu,rele_team,rele_org_name,rele_emp_no,rele_emp_name,reg_date,reg_user) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&chulgo_goods_type&"','"&chulgo_id&"','"&service_no&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"','"&chulgo_type&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&chulgo_emp_no&"','"&chulgo_emp_name&"','"&chulgo_company&"','"&chulgo_bonbu&"','"&chulgo_saupbu&"','"&chulgo_team&"','"&chulgo_org_name&"','"&chulgo_memo&"','"&rele_no&"','"&rele_seq&"','"&rele_stock&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_company&"','"&rele_bonbu&"','"&rele_saupbu&"','"&rele_team&"','"&rele_org_name&"','"&rele_emp_no&"','"&rele_emp_name&"',now(),'"&user_name&"')"

	dbconn.execute(sql)

	j = 0
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			if qty_tab(i) <> 0 then
			   j = j + 1
			cg_goods_seq = right(("00" + cstr(j)),2)
		  	sql="insert into met_chulgo_goods (chulgo_date,chulgo_stock,chulgo_seq,cg_goods_seq,cg_goods_code,cg_goods_type,cg_goods_gubun,cg_standard,cg_goods_name,cg_goods_grade,cg_qty,cg_type,cg_return,rl_service_no,rl_trade_name,rl_trade_dept,reg_date,reg_user,chulgo_stock_company,chulgo_stock_name,in_date,in_no,in_no_seq,in_goods_seq) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&cg_goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"','"&qty_tab(i)&"','"&chulgo_id&"','"&goods_return(i)&"','"&service_no&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"',now(),'"&user_name&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&mvin_in_date&"','"&mvin_in_stock&"','"&mvin_in_seq&"','"&in_goods_seq&"')"
			
			dbconn.execute(sql)
			
' 재고 등록				 
				 sql="select * from met_stock_gmaster where stock_code='"&chulgo_stock&"' and stock_goods_code='"&code_tab(i)&"' and stock_goods_type='"&goods_type(i)&"'"
	             set Rs_jago=dbconn.execute(sql)

                      if not Rs_jago.eof then
						     go_a_qty = Rs_jago("stock_go_qty")
							 JJ_a_qty = Rs_jago("stock_JJ_qty")
							 
							 go_a_qty = go_a_qty + qty_tab(i)
							 JJ_a_qty = JJ_a_qty - qty_tab(i)
							 
	                         sql = "update met_stock_gmaster set stock_go_qty='"&go_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&chulgo_stock&"' and stock_goods_type='"&goods_type(i)&"' and stock_goods_code='"&code_tab(i)&"'"

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
