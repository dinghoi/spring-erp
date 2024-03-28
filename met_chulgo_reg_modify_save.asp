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
	dim goods_seq(20)
	
	dim service_no(20)
	dim trade_name(20)
	dim trade_dept(20)
	dim bigo(20)
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)
	
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50
'2014-01-25 기존에 설치사진 첨부 (종료)

	u_type = abc("u_type")
	
	old_rele_no = abc("old_rele_no")
	old_rele_seq = abc("old_rele_seq")
	old_rele_date = abc("old_rele_date")
	old_att_file = abc("old_att_file")

	rele_date = abc("rele_date")
	rele_id = "고객출고"
	rele_goods_type = abc("rele_goods_type")
	
	rele_stock = abc("rele_stock")
	rele_stock_company = abc("rele_stock_company")
    rele_stock_name = abc("rele_stock_name")
	
    rele_emp_no = abc("rele_emp_no")
    rele_emp_name = abc("rele_emp_name")
	
    rele_company = abc("rele_company")
    rele_bonbu = abc("rele_bonbu")
    rele_saupbu = abc("rele_saupbu")
    rele_team = abc("rele_team")
    rele_org_name = abc("rele_org_name")
    rele_trade_name = ""
	rele_trade_dept = ""
	rele_delivery = ""
    service_acpt = ""
    chulgo_ing = "출고의뢰"
    chulgo_date = abc("chulgo_date")
    chulgo_stock = abc("chulgo_stock")
    chulgo_stock_name = abc("chulgo_stock_name")
	chulgo_stock_company = abc("chulgo_stock_company")
	rele_sign_yn = abc("rele_sign_yn")
	
	if chulgo_date = "" or isnull(chulgo_date) then
		chulgo_date = "0000-00-00"
	end if
	
	rele_memo = abc("chulgo_memo")
	rele_memo = Replace(rele_memo,"'","&quot;")
	
	for i = 1 to 20	
		goods_type(i) = abc("srv_type"&i)
		goods_gubun(i) = abc("goods_gubun"&i)
		code_tab(i) = abc("goods_code"&i)
		goods_name(i) = abc("goods_name"&i)
		goods_standard(i) = abc("goods_standard"&i)
		qty_tab(i) = int(abc("qty"&i))
		goods_grade(i) = abc("goods_grade"&i)
		goods_seq(i) = abc("goods_seq"&i)
		
		service_no(i) = abc("service_no"&i)
		trade_name(i) = abc("trade_name"&i)
		trade_dept(i) = abc("trade_dept"&i)
		bigo(i) = abc("r_bigo"&i)
	next
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	dbconn.BeginTrans

	yymmdd = mid(cstr(rele_date),3,2) + mid(cstr(rele_date),6,2)  + mid(cstr(rele_date),9,2)

    if	u_type = "U" then
	        sql = "delete from met_chulgo_reg where rele_no ='"&old_rele_no&"' and rele_seq='"&old_rele_seq&"' and rele_date='"&old_rele_date&"'"
		    dbconn.execute(sql)
		    sql = "delete from met_chulgo_reg_goods where rele_no ='"&old_rele_no&"' and rele_seq='"&old_rele_seq&"' and rele_date='"&old_rele_date&"'"
		    dbconn.execute(sql)
		
		    rele_no = old_rele_no
			rele_seq = old_rele_seq
			rele_date = old_rele_date
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
			filename = "출고의뢰" + cstr(rele_no) + file_name + "." + fileType
			save_path = path & "\" & filename
		end if
	  else
	  	filename = old_att_file
	end if

	sql="insert into met_chulgo_reg (rele_no,rele_seq,rele_date,rele_id,rele_goods_type,rele_stock,rele_stock_company,rele_stock_name,rele_emp_no,rele_emp_name,rele_company,rele_bonbu,rele_saupbu,rele_team,rele_org_name,rele_trade_name,rele_trade_dept,rele_delivery,service_no,chulgo_ing,chulgo_date,chulgo_stock,chulgo_stock_name,chulgo_stock_company,rele_sign_yn,rele_att_file,rele_memo,reg_date,reg_user) values ('"&rele_no&"','"&rele_seq&"','"&rele_date&"','"&rele_id&"','"&rele_goods_type&"','"&rele_stock&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_emp_no&"','"&rele_emp_name&"','"&rele_company&"','"&rele_bonbu&"','"&rele_saupbu&"','"&rele_team&"','"&rele_org_name&"','"&rele_trade_name&"','"&rele_trade_dept&"','"&rele_delivery&"','"&service_acpt&"','"&chulgo_ing&"','"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_stock_name&"','"&chulgo_stock_company&"','"&rele_sign_yn&"','"&filename&"','"&rele_memo&"',now(),'"&user_name&"')"

	dbconn.execute(sql)

	for i = 1 to 20
		if code_tab(i) = "" or isnull(code_tab(i)) then
			exit for
		  else
			rl_goods_seq = right(("00" + cstr(i)),2)
		  	sql="insert into met_chulgo_reg_goods (rele_no,rele_seq,rele_date,rl_goods_seq,rl_goods_code,rele_stock,rl_goods_type,rl_goods_gubun,rl_standard,rl_goods_name,rl_goods_grade,rl_qty,chulgo_ing,rl_service_no,rl_trade_name,rl_trade_dept,rl_bigo,reg_date,reg_user) values ('"&rele_no&"','"&rele_seq&"','"&rele_date&"','"&rl_goods_seq&"','"&code_tab(i)&"','"&rele_stock&"','"&goods_type(i)&"','"&goods_gubun(i)&"','"&goods_standard(i)&"','"&goods_name(i)&"','"&goods_grade(i)&"','"&qty_tab(i)&"','"&chulgo_ing&"','"&service_no(i)&"','"&trade_name(i)&"','"&trade_dept(i)&"','"&bigo(i)&"',now(),'"&user_name&"')"
			
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
	'response.write"location.replace('meterials_stock_out_mg.asp');"
	response.write"self.opener.location.reload();"	
	response.write"window.close();"	
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
