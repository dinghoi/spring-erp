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
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	
	u_type = request.form("u_type")

	rele_date = request.form("rele_date")
    rele_stock = request.form("rele_stock")
	rele_id = "창고이동"
	rele_goods_type = request.form("u_goods_type")
    rele_stock_company = request.form("rele_stock_company")
    rele_stock_name = request.form("rele_stock_name")
    rele_emp_no = request.form("rele_emp_no")
	rele_emp_name = request.form("rele_emp_name")
    chulgo_type = "출고요청"
    chulgo_rele_date = request.form("chulgo_rele_date")
    chulgo_stock = request.form("chulgo_stock")
    chulgo_stock_name = request.form("chulgo_stock_name")
	chulgo_stock_company = request.form("chulgo_stock_company")
	chulgo_date = "1900-01-01"
	rele_memo = request.form("rele_memo")
	
	for i = 1 to 20	
		goods_type(i) = request.form("srv_type"&i)
		goods_gubun(i) = request.form("goods_gubun"&i)
		code_tab(i) = request.form("goods_code"&i)
		goods_name(i) = request.form("goods_name"&i)
		goods_standard(i) = request.form("goods_standard"&i)
		qty_tab(i) = int(request.form("qty"&i))
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	dbconn.BeginTrans

	yymm = mid(cstr(now()),3,2) + mid(cstr(now()),6,2) 

	sql="select max(rele_seq) as max_seq from met_mv_reg where rele_date = '"&rele_date&"' and rele_stock = '"&rele_stock&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_seq = "01"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		code_seq = right(max_seq,2)
	end if
    rs_max.close()
	
if u_type = "U" then
	   code_last = rele_seq
   else
       code_last = code_seq
end if
	
rele_seq = code_last

	sql="insert into met_mv_reg (rele_date,rele_stock,rele_seq,rele_id,rele_goods_type,rele_stock_company,rele_stock_name,rele_emp_no,rele_emp_name,chulgo_type,chulgo_rele_date,chulgo_stock,chulgo_stock_name,chulgo_stock_company,rele_memo,reg_date,reg_user) values ('"&rele_date&"','"&rele_stock&"','"&rele_seq&"','"&rele_id&"','"&rele_goods_type&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_emp_no&"','"&rele_emp_name&"','"&chulgo_type&"','"&chulgo_rele_date&"','"&chulgo_stock&"','"&chulgo_stock_name&"','"&chulgo_stock_company&"','"&rele_memo&"',now(),'"&user_name&"')"

	dbconn.execute(sql)

	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			bg_seq = right(("00" + cstr(i)),2)
		  	sql="insert into met_mv_reg_goods (rele_date,rele_stock,rele_stock_seq,rele_goods,rele_goods_type,rele_goods_gubun,rele_goods_standard,rele_goods_grade,rele_goods_name,rele_qty,chulgo_date,chulgo_type,reg_date,reg_user) values ('"&rele_date&"','"&rele_stock&"','"&rele_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','A','"&goods_name(i)&"','"&qty_tab(i)&"','"&chulgo_date&"','"&chulgo_type&"',now(),'"&user_name&"')"
			dbconn.execute(sql)
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
