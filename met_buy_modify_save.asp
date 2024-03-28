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
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)

	u_type = abc("u_type")
	old_buy_no = abc("old_buy_no")
	old_buy_seq = abc("old_buy_seq")
	old_buy_date = abc("old_buy_date")
	old_buy_goods_type = abc("old_buy_goods_type")
	old_att_file = abc("old_att_file")
	
	buy_goods_type = abc("buy_goods_type")
	buy_company = abc("buy_company")
	buy_bonbu = abc("emp_bonbu")
	buy_saupbu = abc("buy_saupbu")
	buy_team = abc("emp_team")
	buy_org_code = abc("emp_org_code")
	buy_org_name = abc("emp_org_name")
	buy_emp_no = abc("emp_no")
	buy_emp_name = abc("emp_name")
	
	buy_date = abc("buy_date")
	
	buy_trade_no = abc("trade_no")
	buy_trade_name = abc("trade_name")
	buy_trade_person = abc("trade_person")
	buy_trade_email = abc("trade_email")
	
	'buy_out_method = abc("out_method")
	'buy_out_request_date = abc("out_request_date")
	buy_out_method = ""
	buy_out_request_date = ""
	if buy_out_request_date = "" or isnull(buy_out_request_date) then
		buy_out_request_date = "0000-00-00"
	end if
	
	buy_bill_collect = abc("bill_collect")
	buy_collect_due_date = abc("collect_due_date")
	if buy_collect_due_date = "" or isnull(buy_collect_due_date) then
		buy_collect_due_date = "0000-00-00"
	end if
	
	buy_sign_yn = abc("buy_sign_yn")
	buy_memo = abc("buy_memo")
	buy_memo = Replace(buy_memo,"'","&quot;")

	buy_tot_cost = int(abc("buy_tot_cost"))
	buy_cost_vat = int(abc("buy_tot_cost_vat"))
	buy_price = int(abc("buy_tot_price"))
	
	for i = 1 to 20	
		goods_type(i) = abc("srv_type"&i)
		goods_gubun(i) = abc("goods_gubun"&i)
		code_tab(i) = abc("goods_code"&i)
		
'		response.write(code_tab(i))
		
		goods_name(i) = abc("goods_name"&i)
		goods_standard(i) = abc("goods_standard"&i)
		qty_tab(i) = int(abc("qty"&i))
		buy_cost(i) = int(abc("buy_cost"&i))
		'buy_amt(i) = int(abc("buy_tot"&i))
		buy_amt(i) = qty_tab(i) * buy_cost(i)
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans
	
	yymmdd = mid(cstr(buy_date),3,2) + mid(cstr(buy_date),6,2)  + mid(cstr(buy_date),9,2)
	
	if	u_type = "U" then
		    sql = "delete from met_buy where buy_no ='"&old_buy_no&"' and buy_date='"&old_buy_date&"' and buy_seq='"&old_buy_seq&"'"
		    dbconn.execute(sql)
		    sql = "delete from met_buy_goods where bg_no ='"&old_buy_no&"' and bg_date='"&old_buy_date&"' and buy_seq='"&old_buy_seq&"'"
		    dbconn.execute(sql)
		
		    buy_no = old_buy_no
			buy_seq = old_buy_seq
			buy_date = old_buy_date
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
			filename = "구매" + cstr(buy_no) + file_name + "." + fileType
			save_path = path & "\" & filename
		end if
	  else
	  	filename = old_att_file
	end if
	
	sql="insert into met_buy (buy_no,buy_seq,buy_date,buy_goods_type,buy_company,buy_saupbu,buy_org_code,buy_org_name,buy_emp_no,buy_emp_name,buy_bill_collect,buy_collect_due_date,buy_trade_no,buy_trade_name,buy_trade_person,buy_trade_email,buy_out_method,buy_out_request_date,buy_price,buy_cost,buy_cost_vat,buy_ing,buy_sign_yn,buy_memo,buy_att_file,reg_date,reg_user) values ('"&buy_no&"','"&buy_seq&"','"&buy_date&"','"&buy_goods_type&"','"&buy_company&"','"&buy_saupbu&"','"&buy_org_code&"','"&buy_org_name&"','"&buy_emp_no&"','"&buy_emp_name&"','"&buy_bill_collect&"','"&buy_collect_due_date&"','"&buy_trade_no&"','"&buy_trade_name&"','"&buy_trade_person&"','"&buy_trade_email&"','"&buy_out_method&"','"&buy_out_request_date&"','"&buy_price&"','"&buy_tot_cost&"','"&buy_cost_vat&"','0','"&buy_sign_yn&"','"&buy_memo&"','"&filename&"',now(),'"&user_name&"')"

	dbconn.execute(sql)

	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			bg_seq = right(("00" + cstr(i)),2)
		  	sql="insert into met_buy_goods (bg_no,buy_seq,bg_date,bg_seq,bg_goods_code,bg_goods_type,bg_goods_gubun,bg_goods_name,bg_unit,bg_standard,bg_qty,bg_unit_cost,bg_buy_amt,bg_ing,reg_date,reg_user) values ('"&buy_no&"','"&buy_seq&"','"&buy_date&"','"&bg_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_name(i)&"','','"&goods_standard(i)&"','"&qty_tab(i)&"','"&buy_cost(i)&"','"&buy_amt(i)&"','0',now(),'"&user_name&"')"
			dbconn.execute(sql)
		end if
	next

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
'	response.write"location.replace('meterials_control_mg.asp');"
	response.write"self.opener.location.reload();"	
	response.write"window.close();"	
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
