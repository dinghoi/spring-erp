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
	dim goods_grade(20)
	dim service_no(20)
	dim trade_name(20)
	dim trade_dept(20)
	dim bigo(20)
	
	for i = 1 to 20
        code_tab(i) = ""
	    goods_name(i) = ""
	    goods_type(i) = ""
	    goods_gubun(i) = ""
	    goods_standard(i) = ""
		goods_grade(i) = ""
		qty_tab(i) = 0

		service_no(i) = ""
	    trade_name(i) = ""
	    trade_dept(i) = ""
	    bigo(i) = ""
    next
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	org_name = request.cookies("nkpmg_user")("coo_org_name")
	emp_company = request.cookies("nkpmg_user")("coo_emp_company")
	bonbu = request.cookies("nkpmg_user")("coo_bonbu")
	saupbu = request.cookies("nkpmg_user")("coo_saupbu")
	team = request.cookies("nkpmg_user")("coo_team")
	
	curr_date = mid(cstr(now()),1,10)
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	u_type = abc("u_type")
	
	chulgo_date = abc("chulgo_date")
	chulgo_id = abc("chulgo_id")
	chulgo_goods_type = abc("chulgo_goods_type")
	
	chulgo_stock = abc("chulgo_stock")
    chulgo_stock_name = abc("chulgo_stock_name")
	chulgo_stock_company = abc("chulgo_stock_company")
	
	chulgo_company = emp_company
	chulgo_bonbu = bonbu
	chulgo_saupbu = saupbu
	chulgo_team = team
	chulgo_org_name = org_name
	
	chulgo_ing = "출고완료"
	chulgo_type = "출고완료"
	
	chulgo_emp_no = user_id
	chulgo_emp_name = user_name
	
	chulgo_memo = abc("chulgo_memo")
	chulgo_memo = Replace(chulgo_memo,"'","&quot;")
	
	rele_stock = abc("rele_stock")
	rele_stock_company = abc("rele_stock_company")
    rele_stock_name = abc("rele_stock_name")
	
    rele_emp_no = ""
    rele_emp_name = ""
	
    rele_company = abc("rele_company")
    rele_bonbu = ""
    rele_saupbu = abc("rele_saupbu")
    rele_team = ""
    rele_org_name = ""
    rele_trade_name = ""
	rele_trade_dept = ""
	rele_delivery = ""
    service_acpt = ""
	
	mvin_in_date = abc("mvin_in_date")
	mvin_in_stock = abc("mvin_in_stock")
	mvin_in_seq = abc("mvin_in_seq")
	in_goods_seq = abc("in_goods_seq")
	in_goods_code = abc("in_goods_code")
	
	
	for i = 1 to 20	
'		goods_type(i) = abc("srv_type"&i)
		goods_type(i) = chulgo_goods_type
		goods_gubun(i) = abc("goods_gubun"&i)
		code_tab(i) = abc("goods_code"&i)
		goods_name(i) = abc("goods_name"&i)
		goods_standard(i) = abc("goods_standard"&i)
		qty_tab(i) = int(abc("qty"&i))
		goods_grade(i) = abc("goods_grade"&i)
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_stock = Server.CreateObject("ADODB.Recordset")
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
			filename = "CE출고" + cstr(rele_no) + file_name + "." + fileType
			save_path = path & "\" & filename
		end if
	  else
	  	filename = old_att_file
	end if

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
	
	sql="insert into met_chulgo (chulgo_date,chulgo_stock,chulgo_seq,chulgo_goods_type,chulgo_id,service_no,chulgo_trade_name,chulgo_trade_dept,chulgo_type,chulgo_stock_company,chulgo_stock_name,chulgo_emp_no,chulgo_emp_name,chulgo_company,chulgo_bonbu,chulgo_saupbu,chulgo_team,chulgo_org_name,chulgo_att_file,chulgo_memo,rele_no,rele_seq,rele_stock,rele_stock_company,rele_stock_name,rele_company,rele_bonbu,rele_saupbu,rele_team,rele_org_name,rele_emp_no,rele_emp_name,reg_date,reg_user) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&chulgo_goods_type&"','"&chulgo_id&"','"&service_acpt&"','"&chulgo_trade_name&"','"&chulgo_trade_dept&"','"&chulgo_type&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&chulgo_emp_no&"','"&chulgo_emp_name&"','"&chulgo_company&"','"&chulgo_bonbu&"','"&chulgo_saupbu&"','"&chulgo_team&"','"&chulgo_org_name&"','"&filename&"','"&chulgo_memo&"','"&rele_no&"','"&rele_seq&"','"&rele_stock&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_company&"','"&rele_bonbu&"','"&rele_saupbu&"','"&rele_team&"','"&rele_org_name&"','"&rele_emp_no&"','"&rele_emp_name&"',now(),'"&user_name&"')"

	dbconn.execute(sql)
	
'출고 품목 등록	
	j = 0
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		    if qty_tab(i) <> 0 then
			     j = j + 1
				 cg_goods_seq = right(("00" + cstr(j)),2)
		  	     sql="insert into met_chulgo_goods (chulgo_date,chulgo_stock,chulgo_seq,cg_goods_seq,cg_goods_code,cg_goods_type,cg_goods_gubun,cg_standard,cg_goods_name,cg_goods_grade,rl_qty,cg_qty,cg_type,rl_service_no,rl_trade_name,rl_trade_dept,rl_stock_code,rl_stock_company,rl_stock_name,rl_company,rl_saupbu,rl_bigo,reg_date,reg_user,chulgo_stock_company,chulgo_stock_name,in_date,in_no,in_no_seq,in_goods_seq) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&cg_goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"',0,'"&qty_tab(i)&"','"&chulgo_id&"','"&service_no(i)&"','"&trade_name(i)&"','"&trade_dept(i)&"','"&rele_stock&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_company&"','"&rele_saupbu&"','"&bigo(i)&"',now(),'"&user_name&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&mvin_in_date&"','"&mvin_in_stock&"','"&mvin_in_seq&"','"&in_goods_seq&"')"
				 
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
				 
		'response.write sql
		
		                     dbconn.execute(sql)	  
                       end if				 
			end if
		end if
	next

	
	if filenm <> "" then 
		filenm.save save_path
	end if


'요청창고 입고등록	
mvin_id = "CE출고"
rele_date = chulgo_date


	Sql = "SELECT * FROM met_stock_code where stock_code = '"&rele_stock&"'"
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

	sql="select max(mvin_in_seq) as max_seq from met_mv_in where mvin_in_date = '"&chulgo_date&"' and mvin_in_stock = '"&rele_stock&"'"
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
	
	sql="insert into met_mv_in (mvin_in_date,mvin_in_stock,mvin_in_seq,mvin_id,mvin_goods_type,mvin_stock_company,mvin_stock_name,mvin_emp_no,mvin_emp_name,mvin_company,mvin_bonbu,mvin_saupbu,mvin_team,mvin_org_name,rele_date,rele_stock,rele_seq,rele_no,chulgo_date,chulgo_stock,chulgo_seq,chulgo_stock_company,chulgo_stock_name,chulgo_memo,reg_date,reg_user) values ('"&chulgo_date&"','"&rele_stock&"','"&mvin_in_seq&"','"&mvin_id&"','"&chulgo_goods_type&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_emp_no&"','"&rele_emp_name&"','"&rele_company&"','"&rele_bonbu&"','"&rele_saupbu&"','"&rele_team&"','"&rele_org_name&"','"&rele_date&"','"&rele_stock&"','"&rele_seq&"','"&rele_no&"','"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&chulgo_memo&"',now(),'"&user_name&"')"

	dbconn.execute(sql)

	
'요청창고 입고 품목 등록	
	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
		    if qty_tab(i) <> 0 then
			     goods_seq = right(("00" + cstr(i)),2)
		  	     sql="insert into met_mv_in_goods (mvin_in_date,mvin_in_stock,mvin_in_seq,in_goods_seq,in_goods_code,in_goods_type,in_goods_gubun,in_standard,in_goods_name,in_goods_grade,in_qty,mvin_id,rele_no,rele_seq,reg_date,reg_user) values ('"&chulgo_date&"','"&rele_stock&"','"&mvin_in_seq&"','"&goods_seq&"','"&code_tab(i)&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"','"&qty_tab(i)&"','"&mvin_id&"','"&rele_no&"','"&rele_seq&"',now(),'"&user_name&"')"
			     dbconn.execute(sql)
'			end if

' 재고 등록				 
                 
				stock_goods_level1 = mid(cstr(code_tab(i)),1,3)
                stock_goods_level2 = mid(cstr(code_tab(i)),4,4)
                stock_goods_seq = mid(cstr(code_tab(i)),8,3) 
				 
				 sql="select * from met_stock_gmaster where stock_code='"&rele_stock&"' and stock_goods_type='"&chulgo_goods_type&"' and stock_goods_code='"&code_tab(i)&"'"
	             set Rs_jago=dbconn.execute(sql)

                      if Rs_jago.eof then
                             sql = "insert into met_stock_gmaster(stock_code,stock_goods_type,stock_goods_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_goods_level1,stock_goods_level2,stock_goods_seq,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_in_qty,stock_in_amt,stock_JJ_qty,stock_jj_amt,reg_date,reg_user) values "
	                         sql = sql +	" ('"&rele_stock&"','"&chulgo_goods_type&"','"&code_tab(i)&"','"&stock_level&"','"&stock_name&"','"&stock_company&"','"&stock_bonbu&"','"&stock_saupbu&"','"&stock_team&"','"&stock_goods_level1&"','"&stock_goods_level2&"','"&stock_goods_seq&"','"&goods_gubun(i)&"','"&goods_name(i)&"','"&goods_standard(i)&"','"&goods_grade(i)&"','"&qty_tab(i)&"',0,'"&qty_tab(i)&"',0,now(),'"&user_name&"')"

		'response.write(sql)
		                     dbconn.execute(sql)	 
	                     else
						     in_a_qty = Rs_jago("stock_in_qty")
							 JJ_a_qty = Rs_jago("stock_JJ_qty")
							 
							 in_a_qty = in_a_qty + qty_tab(i)
							 JJ_a_qty = JJ_a_qty + qty_tab(i)
							 
	                         sql = "update met_stock_gmaster set stock_in_qty='"&in_a_qty&"',stock_JJ_qty='"&JJ_a_qty&"',mod_date=now(),mod_user='"&user_name&"' where stock_code='"&rele_stock&"' and stock_goods_type='"&goods_type(i)&"' and stock_goods_code='"&code_tab(i)&"'"

		'response.write sql
		
		                     dbconn.execute(sql)	  
                       end if
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
