<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	dim code_tab(20)
	dim goods_name(20)
	dim goods_type(20)
	dim goods_gubun(20)
	dim goods_standard(20)
	dim qty_tab(20)
	dim amt_tab(20)
	dim goods_grade(20)
	dim excel_file(20)
	dim part_number(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
		goods_grade(i) = ""
		excel_file(i) = ""
		part_number(i) = ""
		qty_tab(i) = 0
		amt_tab(i) = 0
    next
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	
	u_type = abc("u_type")

	chulgo_date = abc("chulgo_date")
    chulgo_stock = abc("chulgo_stock")
	chulgo_id = "NW출고"
	chulgo_goods_type = abc("chulgo_goods_type")
	service_no = abc("service_no")
	chulgo_trade_name = abc("chulgo_trade_name")
	chulgo_trade_dept = abc("chulgo_trade_dept")
	
    chulgo_stock_name = abc("chulgo_stock_name")
	chulgo_stock_company = abc("chulgo_stock_company")
	
	chulgo_emp_no = abc("chulgo_emp_no")
    chulgo_emp_name = abc("chulgo_emp_name")
	chulgo_company = abc("chulgo_company")
    chulgo_bonbu = abc("chulgo_bonbu")
    chulgo_saupbu = abc("chulgo_saupbu")
    chulgo_team = abc("chulgo_team")
    chulgo_org_name = abc("chulgo_org_name")
	
	chulgo_ing = "출고완료"
	chulgo_type = "출고완료"
	
	chulgo_memo = abc("chulgo_memo")
	chulgo_memo = Replace(chulgo_memo,"'","&quot;")
	
	chulgo_price = int(abc("buy_tot_price"))
	chulgo_cost = int(abc("buy_tot_cost"))
	chulgo_cost_vat = int(abc("buy_tot_cost_vat"))
	
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
	
	stin_date = abc("stin_date")
	stin_order_no = abc("stin_order_no")
	stin_order_seq = abc("stin_order_seq")
	stin_goods_seq = abc("stin_goods_seq")
	stin_goods_code = abc("stin_goods_code")
	
	for i = 1 to 20	
'		goods_type(i) = abc("srv_type"&i)
		goods_type(i) = chulgo_goods_type
		goods_gubun(i) = abc("goods_gubun"&i)
		code_tab(i) = abc("goods_code"&i)
		goods_name(i) = abc("goods_name"&i)
		goods_standard(i) = abc("goods_standard"&i)
		part_number(i) = abc("part_number"&i)
		qty_tab(i) = int(abc("qty"&i))
		amt_tab(i) = int(abc("c_amt"&i))
		goods_grade(i) = abc("goods_grade"&i)
	next
	
	set cn = Server.CreateObject("ADODB.Connection")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	Set Rs_mvin = Server.CreateObject("ADODB.Recordset")
	Set Rs_seri = Server.CreateObject("ADODB.Recordset")
	Set Rs_stin = Server.CreateObject("ADODB.Recordset")
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
	
	path_nm = "D:\web\met_upload"
	Set fso=Server.CreateObject("Scripting.FileSystemObject")'
	if Not fso.FolderExists(path_nm) then
		path_nm = fso.CreateFolder(path_nm)
	end if
	Set fso = Nothing
	
	path_name = "/met_upload"
	path = Server.MapPath (path_name)
	
    chulgo_seq = code_seq
    rele_no = yymmdd + chulgo_stock
    rele_seq = chulgo_seq

'입고품목 출고수량 update
    j = 1
	sql = "select * from met_stin_goods where (stin_date = '"&stin_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"') and (stin_goods_seq = '"&stin_goods_seq&"') and (stin_goods_code = '"&stin_goods_code&"')"
	
    set Rs_stin=dbconn.execute(sql)
    if not Rs_stin.eof then
 	       cg_qty = Rs_stin("cg_qty")
		   cg_qty = cg_qty + qty_tab(j)
		   
           sql = "Update met_stin_goods set chulgo_no='"&rele_no&"',chulgo_seq='"&rele_seq&"',cg_qty='"&cg_qty&"' where (stin_date = '"&stin_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"') and (stin_goods_seq = '"&stin_goods_seq&"') and (stin_goods_code = '"&stin_goods_code&"')"
	
	       dbconn.execute(sql)
    end if

'출고등록	

	sql="insert into met_chulgo (chulgo_date,chulgo_stock,chulgo_seq,chulgo_goods_type,chulgo_id,service_no,chulgo_trade_name,chulgo_trade_dept,chulgo_type,chulgo_stock_company,chulgo_stock_name,chulgo_emp_no,chulgo_emp_name,chulgo_company,chulgo_bonbu,chulgo_saupbu,chulgo_team,chulgo_org_name,chulgo_memo,rele_no,rele_seq,rele_stock,rele_stock_company,rele_stock_name,rele_company,rele_bonbu,rele_saupbu,rele_team,rele_org_name,rele_emp_no,rele_emp_name,reg_date,reg_user,chulgo_price,chulgo_cost,chulgo_cost_vat) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&chulgo_goods_type&"','"&chulgo_id&"','"&service_no&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"','"&chulgo_type&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&chulgo_emp_no&"','"&chulgo_emp_name&"','"&chulgo_company&"','"&chulgo_bonbu&"','"&chulgo_saupbu&"','"&chulgo_team&"','"&chulgo_org_name&"','"&chulgo_memo&"','"&rele_no&"','"&rele_seq&"','"&rele_stock&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_company&"','"&rele_bonbu&"','"&rele_saupbu&"','"&rele_team&"','"&rele_org_name&"','"&rele_emp_no&"','"&rele_emp_name&"',now(),'"&user_name&"','"&chulgo_price&"','"&chulgo_cost&"','"&chulgo_cost_vat&"')"

	dbconn.execute(sql)

	j = 0
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			if qty_tab(i) <> 0 then
			   j = j + 1
			
			     unit_danga = amt_tab(i) / qty_tab(i)
' serial_no 저장					 
				 filename = ""
				 Set filenm = abc("excel_att_file"&i)(1)
		         filename = filenm
		         if filenm <> "" then 
			         filename = filenm.safeFileName	
			         file_name = mid(filename,1,inStrRev(filename,".")-1)
			         fileType = mid(filename,inStrRev(filename,".")+1)
			         filename = "NW출고_serial" + cstr(rele_no) + cstr(rele_seq) + file_name + "." + fileType
			         save_path = path & "\" & filename
		         end if
				 
				 if filenm <> "" then 
		            filenm.save save_path
				 
'				    response.write("file_name")
'				    response.write(save_path)
'                    response.End()
			 
				    objFile = save_path
				    cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	                rs.Open "select * from [1:10000]",cn,"0"
				 
				    rowcount=-1
	                xgr = rs.getrows
	                rowcount = ubound(xgr,2)
	                fldcount = rs.fields.count

	                tot_cnt = rowcount + 1
                    if rowcount > -1 then
		               for si=0 to rowcount
			             if xgr(1,si) = "" or isnull(xgr(1,si)) then
				             exit for
			             end if
			             serial_seq = xgr(0,si)
						 serial_no = xgr(1,si)
						 serial_bigo = xgr(4,si)
						 
			             w_cnt = w_cnt + 1
				 
				         sql = "select * from met_goods_serial where goods_code = '"&code_tab(i)&"' and serial_no = '"&serial_no&"' and serial_seq = '"&serial_seq&"'"
						 
		                 set Rs_seri=dbconn.execute(sql)				
		                 if Rs_seri.eof then
							   sql="insert into met_goods_serial (goods_code,serial_no,serial_seq,goods_gubun,goods_name,goods_standard,part_number,chulgo_date,chulgo_no,chulgo_seq,chlugo_amt,chulgo_trade,chulgo_trade_dept,chulgo_bigo,reg_date,reg_user) values ('"&code_tab(i)&"','"&serial_no&"','"&serial_seq&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&goods_standard(i)&"','"&goods_standard(i)&"','"&chulgo_date&"','"&rele_no&"','"&rele_seq&"','"&unit_danga&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"','"&serial_bigo&"',now(),'"&emp_user&"')"
		                       dbconn.execute(sql)
							   
							 else
							   sql = "update met_goods_serial set chulgo_date='"&chulgo_date&"',chulgo_no='"&rele_no&"',chulgo_seq='"&rele_seq&"',chlugo_amt='"&unit_danga&"',chulgo_trade='"&chulgo_trade_name&"',chulgo_trade_dept='"&chulgo_trade_dept&"',chulgo_bigo='"&serial_bigo&"',mod_date=now(),mod_user='"&user_name&"' where goods_code = '"&code_tab(i)&"' and serial_no = '"&serial_no&"' and serial_seq = '"&serial_seq&"'"
							   
							   dbconn.execute(sql)
		                 end if
		               next
	                end if	
                 end if				
			
			cg_goods_seq = right(("00" + cstr(j)),2)
		  	sql="insert into met_chulgo_goods (chulgo_date,chulgo_stock,chulgo_seq,cg_goods_seq,cg_goods_code,cg_goods_type,cg_goods_gubun,cg_standard,cg_goods_name,cg_goods_grade,cg_qty,cg_amt,cg_type,rl_service_no,rl_trade_name,rl_trade_dept,reg_date,reg_user,chulgo_stock_company,chulgo_stock_name,in_date,in_no,in_no_seq,in_goods_seq) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&cg_goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"','"&qty_tab(i)&"','"&amt_tab(i)&"','"&chulgo_id&"','"&service_no&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"',now(),'"&user_name&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&stin_date&"','"&stin_order_no&"','"&stin_order_seq&"','"&stin_goods_seq&"')"
			
			dbconn.execute(sql)
			
' 재고 등록				 
				 sql="select * from met_stock_gmaster where stock_code='"&chulgo_stock&"' and stock_goods_code='"&code_tab(i)&"' and stock_goods_type='"&goods_type(i)&"'"
	             set Rs_jago=dbconn.execute(sql)

                      if not Rs_jago.eof then
						     go_a_qty = Rs_jago("stock_go_qty")
							 JJ_a_qty = Rs_jago("stock_JJ_qty")
							 
							 go_a_amt = Rs_jago("stock_go_amt")
							 JJ_a_amt = Rs_jago("stock_jj_amt")
							 
							 go_a_qty = go_a_qty + qty_tab(i)
							 JJ_a_qty = JJ_a_qty - qty_tab(i)
							 
							 go_a_amt = go_a_amt + amt_tab(i)
							 JJ_a_amt = JJ_a_amt - amt_tab(i)
							 
	                         sql = "update met_stock_gmaster set stock_go_qty='"&go_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',stock_go_amt='"&go_a_amt&"',stock_jj_amt='"&JJ_a_amt&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&chulgo_stock&"' and stock_goods_type='"&goods_type(i)&"' and stock_goods_code='"&code_tab(i)&"'"

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
