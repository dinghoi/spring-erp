<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim abc,filenm
	dim code_tab(20)
	dim qty_tab(20)
	dim order_qty_tab(20)
	dim srv_tab(20)
	dim pummok_tab(20)
	dim standard_tab(20)
	dim buy_tab(20)
	dim sales_tab(20)
	dim order_tab(20)
	dim margin_tab(20)
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)

	u_type = abc("u_type")
	slip_id = abc("slip_id")
	slip_no = abc("slip_no")
	slip_seq = abc("slip_seq")
'	response.write(slip_id)
'	response.write(slip_no)
'	response.write(slip_seq)
	sales_company = abc("sales_company")
	sales_saupbu = abc("sales_saupbu")
	sales_bonbu = abc("sales_saupbu")
	sales_team = abc("sales_saupbu")
	sales_org_name = abc("sales_saupbu")
'	trade_no = abc("trade_no")
	trade_no = replace(abc("trade_no"),"-","")
	trade_name = abc("trade_name")
	trade_person = abc("trade_person")
	trade_email = abc("trade_email")
	out_method = abc("out_method")
	out_request_date = abc("out_request_date")
	sales_date = abc("sales_date")
	sales_yn = abc("sales_yn")
	bill_due_date = abc("bill_due_date")
	bill_issue_yn = abc("bill_issue_yn")
	bill_issue_date = abc("bill_issue_date")
	bill_collect = abc("bill_collect")
	collect_stat = abc("collect_stat")
	collect_date = abc("collect_date")
	collect_due_date = abc("collect_due_date")
	slip_memo = abc("slip_memo")
	slip_memo = Replace(slip_memo,"'","&quot;")

	if collect_date = "" or isnull(collect_date) then
		collect_date = "0000-00-00"
	end if
	if bill_issue_date = "" or isnull(bill_issue_date) then
		bill_issue_date = "0000-00-00"
	end if
	if collect_due_date = "" or isnull(collect_due_date) then
		collect_due_date = "0000-00-00"
	end if
	sales_tot_cost = int(abc("sales_tot_cost"))
	sales_tot_cost_vat = int(sales_tot_cost * 0.1)		
	sales_tot_price = sales_tot_cost + sales_tot_cost_vat
	buy_tot_cost = int(abc("buy_tot_cost"))
	buy_tot_cost_vat = int(buy_tot_cost * 0.1)		
	buy_tot_price = buy_tot_cost + buy_tot_cost_vat
	margin_tot_cost = int(abc("margin_tot_cost"))
	old_att_file = abc("old_att_file")
	sign_yn = abc("sign_yn")
'	pg_name = abc("pg_name")
	
	for i = 1 to 20	
		code_tab(i) = abc("pummok_code"&i)
		srv_tab(i) = abc("srv_type"&i)
		pummok_tab(i) = abc("pummok"&i)
		standard_tab(i) = abc("standard"&i)
		qty_tab(i) = int(abc("qty"&i))
		order_qty_tab(i) = int(abc("order_qty"&i))
		buy_tab(i) = int(abc("buy_cost"&i))
		sales_tab(i) = int(abc("sales_cost"&i))
		order_tab(i) = int(abc("order_cost"&i))
'		margin_tab(i) = int(abc("margin_cost"&i))		
		margin_tab(i) = sales_tab(i) - buy_tab(i)
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	Set filenm = abc("att_file")(1)
	if filenm <> "" then
		path_nm = "D:\web\sales_file"
		Set fso=Server.CreateObject("Scripting.FileSystemObject")'
		if Not fso.FolderExists(path_nm) then
			path_nm = fso.CreateFolder(path_nm)
		end if
		Set fso = Nothing
	
		path_name = "/sales_file"
		path = Server.MapPath (path_name)
	
		Set filenm = abc("att_file")(1)
		filename = filenm
		if filenm <> "" then 
			filename = filenm.safeFileName	
			file_name = mid(filename,1,inStrRev(filename,".")-1)
			fileType = mid(filename,inStrRev(filename,".")+1)
			filename = cstr(slip_id) + cstr(slip_no) + cstr(slip_seq) + file_name + "." + fileType
			save_path = path & "\" & filename
		end if
	  else
	  	filename = old_att_file
	end if
	
' 수주전표 생성

	sql="select max(slip_seq) as max_seq from sales_slip where slip_id='2' and slip_no='"&slip_no&"'"
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_seq"))  then
		order_slip_seq = "00"
	  else
		max_seq = "0" + cstr((int(rs("max_seq")) + 1))
		order_slip_seq = right(max_seq,2)
	end if

	sql="insert into sales_slip (slip_id,slip_no,slip_seq,sales_company,sales_bonbu,sales_saupbu,sales_team,sales_org_name,emp_no,emp_name,emp_company,bonbu,saupbu,team,org_name,trade_no,trade_name,trade_person,trade_email,out_method,out_request_date,sales_date,sales_yn,bill_due_date,bill_issue_yn,bill_issue_date,bill_collect,collect_due_date,collect_stat,collect_date,slip_memo,sales_price,sales_cost,sales_cost_vat,buy_price,buy_cost,buy_cost_vat,margin_cost,att_file,reg_emp_no,reg_name,reg_date) values ('2','"&slip_no&"','"&order_slip_seq&"','"&sales_company&"','"&sales_bonbu&"','"&sales_saupbu&"','"&sales_team&"','"&sales_org_name&"','"&user_id&"','"&user_name&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&trade_no&"','"&trade_name&"','"&trade_person&"','"&trade_email&"','"&out_method&"','"&out_request_date&"','"&sales_date&"','"&sales_yn&"','"&bill_due_date&"','"&bill_issue_yn&"','"&bill_issue_date&"','"&bill_collect&"','"&collect_due_date&"','"&collect_stat&"','"&collect_date&"','"&slip_memo&"',"&sales_tot_price&","&sales_tot_cost&","&sales_tot_cost_vat&","&buy_tot_price&","&buy_tot_cost&","&buy_tot_cost_vat&","&margin_tot_cost&",'"&filename&"','"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)

	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			goods_seq = right(("00" + cstr(i)),2)
		  	sql="insert into sales_slip_detail (slip_id,slip_no,slip_seq,goods_seq,goods_code,srv_type,pummok,standard,qty,buy_cost,sales_cost,margin_cost) values ('2','"&slip_no&"','"&order_slip_seq&"','"&goods_seq&"','"&code_tab(i)&"','"&srv_tab(i)&"','"&pummok_tab(i)&"','"&standard_tab(i)&"',"&order_qty_tab(i)&","&buy_tab(i)&","&order_tab(i)&","&margin_tab(i)&")"
			dbconn.execute(sql)
		end if
	next

' 기존 대기전표에 수주된 금액 업데이트

	order_end_sw = "Y"
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			goods_seq = right(("00" + cstr(i)),2)

			sql = "select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"' and goods_code = '"&code_tab(i)&"' and qty = "&int(qty_tab(i))
			set rs=dbconn.execute(sql)
			sum_order_qty = int(rs("order_qty")) + int(order_qty_tab(i))
			if rs("qty") <> sum_order_qty then
				order_end_sw = "N"
			end if

			sql = "Update sales_slip_detail set order_qty ="&sum_order_qty&" where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"' and goods_code = '"&code_tab(i)&"' and qty = "&int(qty_tab(i))
			dbconn.execute(sql)
		end if
	next

	sql = "select * from sales_slip where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
	set rs=dbconn.execute(sql)
	sum_order_price = int(rs("order_price")) + int(sales_tot_price)
	sum_order_cost = int(rs("order_cost")) + int(sales_tot_cost)
	sum_order_cost_vat = int(rs("order_cost_vat")) + int(sales_tot_cost_vat)
	if sum_order_price = rs("sales_price") or order_end_sw = "Y" then
		slip_stat = "3"
	  else
	  	slip_stat = "2"
	end if

	sql = "Update sales_slip set slip_stat ='"&slip_stat&"', order_price ="&sum_order_price&", order_cost ="&sum_order_cost&", order_cost_vat ="&sum_order_cost_vat&"  where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
	dbconn.execute(sql)

	sql = "Update sales_slip set slip_stat ='"&slip_stat&"' where slip_no = '"&slip_no&"' and slip_id = '2' and slip_seq = '"&order_slip_seq&"'"
	dbconn.execute(sql)

	if filenm <> "" then 
		filenm.save save_path
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	if u_type = "U" then
		response.write"opener.document.frm.submit();"
		response.write"window.close();"		
	  else	
		response.write"location.replace('sales_slip_ing_mg.asp');"		
	end if
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
