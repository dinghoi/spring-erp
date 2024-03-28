<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next
	
    curr_date = mid(cstr(now()),1,10)
	
	u_type = request.form("u_type")
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	'goods_type = request.form("goods_type")
	goods_type = ""
	goods_code = request.form("goods_code")
	goods_name = request.form("goods_name")
	goods_level1 = request.form("goods_level1")
	goods_level2 = request.form("goods_level2")
	goods_seq = request.form("goods_seq")
	goods_grade = request.form("goods_grade")
	goods_gubun = request.form("goods_gubun")
	goods_model = request.form("goods_model")
	goods_serial_no = request.form("goods_serial_no")
	goods_standard = request.form("goods_standard")
	part_number = request.form("part_number")
	po_number = ""
	goods_group = request.form("goods_group")
	goods_date = request.form("goods_date")
	goods_tax_id = request.form("goods_tax_id")
	goods_stock_In_type = request.form("goods_stock_In_type")
	goods_security_yn = request.form("goods_security_yn")
	goods_security_cnt =int(request.form("goods_security_cnt"))
	goods_used_sw = request.form("goods_used_sw")
	goods_comment = request.form("goods_comment")
	goods_comment = Replace(goods_comment,"'","&quot;")
	
	goods_bigo = request.form("goods_bigo")
	goods_end_date = request.form("goods_end_date")
	
    if goods_date = "" or isnull(goods_date) then
	   goods_date = curr_date
	end if
	
	if goods_end_date = "" or isnull(goods_end_date) then
	   goods_end_date = "0000-00-00"
	end if

    set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_max = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

    sql="select max(goods_seq) as max_seq from met_goods_code where goods_level1 = '"&goods_level1&"' and goods_level2 = '"&goods_level2&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_seq = "001"
	  else
		max_seq = "000" + cstr((int(rs_max("max_seq")) + 1))
		code_seq = right(max_seq,3)
	end if
    rs_max.close()
	
if u_type = "U" then
	   code_last = goods_code
   else
       code_last = goods_level1 + goods_level2 + code_seq
	   goods_seq = code_seq
end if
	
goods_code = code_last

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "update met_goods_code set goods_type='"&goods_type&"',goods_gubun='"&goods_gubun&"',goods_grade='"&goods_grade&"',goods_model='"&goods_model&"',goods_group='"&goods_group&"',goods_serial_no='"&goods_serial_no&"',goods_name='"&goods_name&"',goods_standard='"&goods_standard&"',goods_date='"&goods_date&"',goods_used_sw='"&goods_used_sw&"',goods_end_date='"&goods_end_date&"',goods_tax_id='"&goods_tax_id&"',goods_stock_In_type='"&goods_stock_In_type&"',goods_security_yn='"&goods_security_yn&"',goods_security_cnt='"&goods_security_cnt&"',goods_comment='"&goods_comment&"',goods_bigo='"&goods_bigo&"',part_number='"&part_number&"',po_number='"&po_number&"',mod_date=now(),mod_user='"&user_name&"' where goods_code = '"&goods_code&"'"

		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into met_goods_code (goods_code,goods_type,goods_level1,goods_level2,goods_seq,goods_grade,goods_gubun,goods_model,goods_group,goods_serial_no,goods_name,goods_standard,goods_date,goods_used_sw,goods_end_date,goods_tax_id,goods_stock_In_type,goods_security_yn,goods_security_cnt,goods_comment,goods_bigo,part_number,po_number"
		sql = sql + ",reg_date,reg_user) values "
		sql = sql + " ('"&goods_code&"','"&goods_type&"','"&goods_level1&"','"&goods_level2&"','"&goods_seq&"','"&goods_grade&"','"&goods_gubun&"','"&goods_model&"','"&goods_group&"','"&goods_serial_no&"','"&goods_name&"','"&goods_standard&"','"&goods_date&"','"&goods_used_sw&"','"&goods_end_date&"','"&goods_tax_id&"','"&goods_stock_In_type&"','"&goods_security_yn&"','"&goods_security_cnt&"','"&goods_comment&"','"&goods_bigo&"','"&part_number&"','"&po_number&"',now(),'"&user_name&"')"
		
		'response.write sql
		
		dbconn.execute(sql)
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
	response.write"self.opener.location.reload();"		
	response.write"window.close();"			
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
