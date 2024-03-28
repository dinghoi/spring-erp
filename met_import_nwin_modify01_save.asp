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
	dim part_number(20)
	dim qty_tab(20)
	dim buy_cost(20)
	dim buy_amt(20)
	dim d_cost(20)
	dim ex_cost(20)
	dim w_won(20)
	dim excel_file(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    part_number(i) = ""
		excel_file(i) = ""
		qty_tab(i) = 0
	    buy_cost(i) = 0
	    buy_amt(i) = 0
		d_cost(i) = 0
		ex_cost(i) = 0
		w_won(i) = 0
    next
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)

	u_type = abc("u_type")
	old_stin_in_date = abc("old_stin_in_date")
	old_stin_order_no = abc("old_stin_order_no")
	old_stin_order_seq = abc("old_stin_order_seq")
	old_stin_goods_type = abc("old_stin_goods_type")
	old_stin_att_file = abc("old_stin_att_file")
	
	buy_goods_type = abc("buy_goods_type")
	buy_company = abc("buy_company")
	buy_bonbu = abc("emp_bonbu")
	buy_saupbu = abc("buy_saupbu")
	
'	stin_id = abc("stin_id")
    stin_id = "외자구매"
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
	
	po_date = abc("po_date")
	po_number = abc("po_number")
	park_bl = abc("park_bl")


	stin_price = int(abc("buy_tot_price"))
	stin_cost = int(abc("buy_tot_cost"))
	stin_cost_vat = int(abc("buy_tot_cost_vat"))
	
	won_ex = 0
'	won_ex = abc("exchan_rate")
	won_ex = cint(abc("exchan_rate"))
'	exchan_rate = formatnumber(abc("exchan_rate"),2)
	tong_cost = int(abc("tong_cost"))
	stock_cost = int(abc("stock_cost"))
	trans_cost = int(abc("trans_cost"))
	air_cost = int(abc("air_cost"))
	inland_cost = int(abc("inland_cost"))
	
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
		part_number(i) = abc("part_number"&i)
		if code_tab(i) = "" or isnull(code_tab(i)) then
			  exit for
		   else
		      qty_tab(i) = int(abc("qty"&i))
		      d_cost(i) = abc("d_cost"&i)
		      ex_cost(i) = int(abc("ex_cost"&i))
		      w_won(i) = abc("w_won"&i)
		      buy_cost(i) = int(abc("buy_cost"&i))
		      buy_amt(i) = int(abc("buy_tot"&i))
		      'buy_amt(i) = qty_tab(i) * buy_cost(i)
		end if
	next
	
	set cn = Server.CreateObject("ADODB.Connection")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
	Set Rs_stin = Server.CreateObject("ADODB.Recordset")
	Set Rs_goods = Server.CreateObject("ADODB.Recordset")
	Set Rs_seri = Server.CreateObject("ADODB.Recordset")
	Set Rs_jago = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

	dbconn.BeginTrans
	
	yymmdd = mid(cstr(stin_in_date),3,2) + mid(cstr(stin_in_date),6,2)  + mid(cstr(stin_in_date),9,2)
	
	if	u_type = "U" then

' 재고조정
        Sql="select * from met_stin_goods where (stin_date = '"&old_stin_in_date&"') and (stin_order_no = '"&old_stin_order_no&"') and (stin_order_seq = '"&old_stin_order_seq&"')"
	    Set Rs_stin=DbConn.Execute(Sql)
		do until Rs_stin.eof
            mod_stock_code = Rs_stin("stin_stock_code")
			mod_goods_type = Rs_stin("stin_goods_type")
			mod_goods_code = Rs_stin("stin_goods_code")
			mod_stock_company = Rs_stin("stin_stock_company")
			
			mod_stin_qty = Rs_stin("stin_qty")
			mod_stin_amt = Rs_stin("stin_amt")
     
			sql="select * from met_stock_gmaster where stock_code='"&mod_stock_code&"' and stock_goods_code='"&mod_goods_code&"' and stock_goods_type='"&mod_goods_type&"'"
	        set Rs_jago=dbconn.execute(sql)

            if not Rs_jago.eof then
			       in_a_qty = Rs_jago("stock_in_qty")
				   in_a_amt = Rs_jago("stock_in_amt")
				   jj_a_qty = Rs_jago("stock_JJ_qty")
				   jj_a_amt = Rs_jago("stock_jj_amt")
							 
				   in_a_qty = in_a_qty - mod_stin_qty
				   in_a_amt = in_a_amt - mod_stin_amt
				   jj_a_qty = jj_a_qty - mod_stin_qty
				   jj_a_amt = jj_a_amt - mod_stin_amt
							 
	               sql = "update met_stock_gmaster set stock_in_qty='"&in_a_qty&"',stock_in_amt='"&in_a_amt&"',stock_JJ_qty='"&JJ_a_qty&"',stock_jj_amt='"&jj_a_amt&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&mod_stock_code&"' and stock_goods_type='"&mod_goods_type&"' and stock_goods_code='"&mod_goods_code&"'"

		          'response.write sql
		
		           dbconn.execute(sql)	  
            end if	 
	 
		    Rs_stin.movenext()
	    loop
        Rs_stin.close()


		sql = "delete from met_stin where stin_in_date ='"&old_stin_in_date&"' and stin_order_no='"&old_stin_order_no&"' and stin_order_seq='"&old_stin_order_seq&"'"
		dbconn.execute(sql)
		sql = "delete from met_stin_goods where stin_date ='"&old_stin_in_date&"' and stin_order_no='"&old_stin_order_no&"' and stin_order_seq='"&old_stin_order_seq&"'"
		dbconn.execute(sql)
' serial no삭제		
		sql = "delete from met_goods_serial where in_date ='"&old_stin_in_date&"' and in_order_no='"&old_stin_order_no&"' and in_order_seq='"&old_stin_order_seq&"'"
		dbconn.execute(sql)
		
		stin_in_date = old_stin_in_date
		stin_order_no = old_stin_order_no
		stin_order_seq = old_stin_order_seq
		buy_no = stin_order_no
		buy_seq = stin_order_seq
	end if
	
	path_nm = "D:\web\met_upload"
	Set fso=Server.CreateObject("Scripting.FileSystemObject")'
	if Not fso.FolderExists(path_nm) then
		path_nm = fso.CreateFolder(path_nm)
	end if
	Set fso = Nothing
	
	path_name = "/met_upload"
	path = Server.MapPath (path_name)
	
	Set filenm1 = abc("att_file")(1)
	if filenm1 <> "" then
		Set filenm1 = abc("att_file")(1)
		filename1 = filenm1
		if filenm1 <> "" then 
			filename1 = filenm1.safeFileName	
			file_name1 = mid(filename1,1,inStrRev(filename1,".")-1)
			fileType1 = mid(filename1,inStrRev(filename1,".")+1)
			filename1 = "외자입고" + cstr(buy_no) + file_name1 + "." + fileType1
			save_path1 = path & "\" & filename1
		end if
	  else
	  	filename1 = old_att_file
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
	
	sql="insert into met_stin (stin_in_date,stin_order_no,stin_order_seq,stin_order_date,stin_buy_no,stin_buy_seq,stin_buy_date,stin_buy_company,stin_buy_bonbu,stin_buy_saupbu,stin_goods_type,stin_trade_no,stin_trade_name,stin_trade_person,stin_trade_email,stin_stock_company,stin_stock_code,stin_stock_name,stin_id,stin_type,stin_bill_collect,stin_collect_due_date,stin_price,stin_cost,stin_cost_vat,stin_company,stin_org_name,stin_emp_no,stin_emp_name,stin_att_file,stin_memo,reg_date,reg_user,tong_cost,stock_cost,trans_cost,air_cost,inland_cost,po_date,po_number,park_bl,won_ex) values ('"&stin_in_date&"','"&stin_order_no&"','"&stin_order_seq&"','"&stin_order_date&"','','','"&stin_buy_date&"','"&buy_company&"','"&buy_bonbu&"','"&buy_saupbu&"','"&buy_goods_type&"','"&order_trade_no&"','"&order_trade_name&"','"&order_trade_person&"','"&order_trade_email&"','"&stin_stock_company&"','"&stin_stock_code&"','"&stin_stock_name&"','"&stin_id&"','"&stin_type&"','','"&stin_collect_due_date&"','"&stin_price&"','"&stin_cost&"','"&stin_cost_vat&"','"&stin_company&"','"&stin_org_name&"','"&stin_emp_no&"','"&stin_emp_name&"','"&filename&"','"&buy_memo&"',now(),'"&user_name&"','"&tong_cost&"','"&stock_cost&"','"&trans_cost&"','"&air_cost&"','"&inland_cost&"','"&po_date&"','"&po_number&"','"&park_bl&"','"&w_won(1)&"')"

	dbconn.execute(sql)
	
'입고 품목 등록	
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		         unit_wonga = buy_cost(i) + ex_cost(i)
' serial_no 저장					 
				 filename = ""
				 Set filenm = abc("excel_att_file"&i)(1)
		         filename = filenm
		         if filenm <> "" then 
			         filename = filenm.safeFileName	
			         file_name = mid(filename,1,inStrRev(filename,".")-1)
			         fileType = mid(filename,inStrRev(filename,".")+1)
			         filename = "외자입고_serial" + cstr(buy_no) + file_name + "." + fileType
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
						 
			             w_cnt = w_cnt + 1
				 
				         sql = "select * from met_goods_serial where goods_code = '"&code_tab(i)&"' and serial_no = '"&serial_no&"' and serial_seq = '"&serial_seq&"'"
		                 set Rs_seri=dbconn.execute(sql)				
		                 if Rs_seri.eof or Rs_seri.bof then
			                   sql="insert into met_goods_serial (goods_code,serial_no,serial_seq,goods_gubun,goods_name,part_number,po_date,po_number,in_date,in_order_no,in_order_seq,d_cost,w_won,won_cost,ex_cost,unit_wonga,reg_date,reg_user) values ('"&code_tab(i)&"','"&serial_no&"','"&serial_seq&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&part_number(i)&"','"&po_date&"','"&po_number&"','"&stin_in_date&"','"&buy_no&"','"&buy_seq&"','"&d_cost(i)&"','"&w_won(i)&"','"&buy_cost(i)&"','"&ex_cost(i)&"','"&unit_wonga&"',now(),'"&emp_user&"')"
		                       dbconn.execute(sql)
		                 end if
		               next
	                end if	
                 end if		
'입고품목 등록
			     goods_seq = right(("00" + cstr(i)),2)
		  	     sql="insert into met_stin_goods (stin_date,stin_order_no,stin_order_seq,stin_goods_code,stin_goods_seq,stin_goods_type,stin_goods_gubun,stin_goods_name,stin_standard,stin_unit_cost,stin_qty,stin_amt,stin_id,stin_type,stin_stock_company,stin_stock_code,stin_stock_name,reg_date,reg_user,d_cost,w_won,ex_cost,part_number,excel_att_file) values ('"&stin_in_date&"','"&buy_no&"','"&buy_seq&"','"&code_tab(i)&"','"&goods_seq&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&part_number(i)&"','"&unit_wonga&"','"&qty_tab(i)&"','"&buy_amt(i)&"','"&stin_id&"','"&stin_type&"','"&stin_stock_company&"','"&stin_stock_code&"','"&stin_stock_name&"',now(),'"&user_name&"','"&d_cost(i)&"','"&w_won(i)&"','"&ex_cost(i)&"','"&part_number(i)&"','"&filename&"')"
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
	                         sql = sql +	" ('"&stin_stock_code&"','"&goods_type(i)&"','"&code_tab(i)&"','"&stock_level&"','"&stock_name&"','"&stock_company&"','"&stock_bonbu&"','"&stock_saupbu&"','"&stock_team&"','"&stock_goods_level1&"','"&stock_goods_level2&"','"&stock_goods_seq&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&part_number(i)&"','"&goods_grade&"','"&qty_tab(i)&"','"&buy_amt(i)&"','"&qty_tab(i)&"','"&buy_amt(i)&"',now(),'"&user_name&"')"

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
    if filenm1 <> "" then 
		filenm1.save save_path1
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
'	response.write"location.replace('met_stock_nwin_report01.asp');"
	response.write"self.opener.location.reload();"	
	response.write"window.close();"	
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
