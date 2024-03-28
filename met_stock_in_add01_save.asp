<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc,filenm
	dim code_tab(20)
	dim goods_name(20)
	dim goods_type(20)
	dim goods_gubun(20)
	dim goods_standard(20)
	dim qty_tab(20)
	dim buy_cost(20)
	dim buy_amt(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
		qty_tab(i) = 0
	    buy_cost(i) = 0
	    buy_amt(i) = 0
    next
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)

	u_type = abc("u_type")
	
	buy_goods_type = abc("buy_goods_type")
	buy_company = abc("buy_company")
	buy_bonbu = abc("emp_bonbu")
	buy_saupbu = abc("buy_saupbu")
	
'	stin_id = abc("stin_id")
    stin_id = "구매"
	stin_in_date = abc("stin_in_date")
	stin_stock_company = abc("stin_stock_company")
	stin_stock_code = abc("stin_stock_code")
    stin_stock_name = abc("stin_stock_name")
	
	stin_emp_no = abc("emp_no")
    stin_emp_name = abc("emp_name")
	stin_company = abc("emp_company")
    stin_org_name = abc("emp_org_name")
	
	'order_trade_no = abc("trade_no")
	order_trade_no = replace(abc("trade_no"),"-","")
	order_trade_name = abc("trade_name")
	order_trade_person = abc("trade_person")
	order_trade_email = abc("trade_email")

	buy_memo = abc("buy_memo")
	buy_memo = Replace(buy_memo,"'","&quot;")

	stin_price = int(abc("buy_tot_price"))
	stin_cost = int(abc("buy_tot_cost"))
	stin_cost_vat = int(abc("buy_tot_cost_vat"))
	
	buy_memo = abc("buy_memo")
	buy_memo = Replace(buy_memo,"'","&quot;")
	
	stin_type = "정상"
	
	stin_collect_due_date = "0000-00-00"
	stin_order_date = "0000-00-00"
	stin_buy_date = "0000-00-00"
	stin_cal_date = "0000-00-00"
		
	for i = 1 to 20	
'		goods_type(i) = abc("srv_type"&i)
		goods_type(i) = buy_goods_type
		goods_gubun(i) = abc("goods_gubun"&i)
		code_tab(i) = abc("goods_code"&i)
		goods_name(i) = abc("goods_name"&i)
		goods_standard(i) = abc("goods_standard"&i)
		qty_tab(i) = int(abc("qty"&i))
		buy_cost(i) = int(abc("buy_cost"&i))
		'buy_tab(i) = int(abc("buy_tot"&i))
		buy_amt(i) = qty_tab(i) * buy_cost(i)
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
	Set Rs_goods = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans
	
	yymmdd = mid(cstr(stin_in_date),3,2) + mid(cstr(stin_in_date),6,2)  + mid(cstr(stin_in_date),9,2)
	
	buy_seq = "00"
	sql="select max(stin_order_no) as max_no from met_stin where stin_in_date = '"&stin_in_date&"'"
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_no"))  then
	         buy_no = yymmdd + "001"
	    else
	         max_seq = "00" + cstr((int(right(rs("max_no"),3)) + 1))
	         buy_no = yymmdd + cstr(right(max_seq,3))
	end if
	
	Set filenm = abc("att_file")(1)
	if filenm <> "" then
		path_nm = "D:\web\met_upload"
		Set fso=Server.CreateObject("Scripting.FileSystemObject")'
		if Not fso.FolderExists(path_nm) then
			path_nm = fso.CreateFolder(path_nm)
		end if
		Set fso = Nothing
	
		path_name = "/met_upload"
		path = Server.MapPath (path_name)
	
		Set filenm = abc("att_file")(1)
		filename = filenm
		if filenm <> "" then 
			filename = filenm.safeFileName	
			file_name = mid(filename,1,inStrRev(filename,".")-1)
			fileType = mid(filename,inStrRev(filename,".")+1)
			filename = "입고" + cstr(buy_no) + file_name + "." + fileType
			save_path = path & "\" & filename
		end if
	  else
	  	filename = old_att_file
	end if

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

'입고등록	
	
	sql="insert into met_stin (stin_in_date,stin_order_no,stin_order_seq,stin_order_date,stin_buy_no,stin_buy_seq,stin_buy_date,stin_buy_company,stin_buy_bonbu,stin_buy_saupbu,stin_goods_type,stin_trade_no,stin_trade_name,stin_trade_person,stin_trade_email,stin_stock_company,stin_stock_code,stin_stock_name,stin_id,stin_type,stin_bill_collect,stin_collect_due_date,stin_price,stin_cost,stin_cost_vat,stin_company,stin_org_name,stin_emp_no,stin_emp_name,stin_att_file,stin_memo,reg_date,reg_user) values ('"&stin_in_date&"','"&buy_no&"','"&buy_seq&"','"&stin_order_date&"','','','"&stin_buy_date&"','"&buy_company&"','"&buy_bonbu&"','"&buy_saupbu&"','"&buy_goods_type&"','"&order_trade_no&"','"&order_trade_name&"','"&order_trade_person&"','"&order_trade_email&"','"&stin_stock_company&"','"&stin_stock_code&"','"&stin_stock_name&"','"&stin_id&"','"&stin_type&"','','"&stin_collect_due_date&"','"&stin_price&"','"&stin_cost&"','"&stin_cost_vat&"','"&stin_company&"','"&stin_org_name&"','"&stin_emp_no&"','"&stin_emp_name&"','"&filename&"','"&buy_memo&"',now(),'"&user_name&"')"

	dbconn.execute(sql)
	
'입고 품목 등록	
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			     goods_seq = right(("00" + cstr(i)),2)
		  	     sql="insert into met_stin_goods (stin_date,stin_order_no,stin_order_seq,stin_goods_code,stin_goods_seq,stin_goods_type,stin_goods_gubun,stin_goods_name,stin_standard,stin_unit_cost,stin_qty,stin_amt,stin_id,stin_type,stin_stock_company,stin_stock_code,stin_stock_name,reg_date,reg_user) values ('"&stin_in_date&"','"&buy_no&"','"&buy_seq&"','"&code_tab(i)&"','"&goods_seq&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&goods_standard(i)&"','"&buy_cost(i)&"','"&qty_tab(i)&"','"&buy_amt(i)&"','"&stin_id&"','"&stin_type&"','"&stin_stock_company&"','"&stin_stock_code&"','"&stin_stock_name&"',now(),'"&user_name&"')"
			     dbconn.execute(sql)
				 
' 재고 등록				 
                 Sql = "SELECT * FROM met_goods_code where goods_code = '"&code_tab(i)&"'"
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

'첨부파일
    if filenm <> "" then 
		filenm.save save_path
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"location.replace('met_stock_in_report01.asp');"
	response.write"self.opener.location.reload();"	
	response.write"window.close();"	
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
