<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	trade_code = request.form("trade_code")
	trade_no1 = request.form("trade_no1")
	trade_no2 = request.form("trade_no2")
	trade_no3 = request.form("trade_no3")
	trade_no = cstr(trade_no1) + cstr(trade_no2) + cstr(trade_no3)
	old_trade_no = request.form("old_trade_no")
	trade_name = request.form("trade_name")
	trade_id = request.form("trade_id")
	sales_type = request.form("sales_type")
	trade_owner = request.form("trade_owner")
	trade_addr = request.form("trade_addr")	
	trade_uptae = request.form("trade_uptae")	
	trade_upjong = request.form("trade_upjong")
	trade_tel = request.form("trade_tel")	
	trade_fax = request.form("trade_fax")
	bill_trade_code = request.form("bill_trade_code")
	bill_trade_name = request.form("bill_trade_name")
	group_name = request.form("group_name")
	use_sw = request.form("use_sw")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "Update trade set trade_no='"&trade_no&"',trade_name ='"&trade_name&"',bill_trade_code='"&bill_trade_code&"',bill_trade_name='"&bill_trade_name&"',trade_id ='"&trade_id&"',sales_type='"&sales_type&"',trade_owner='"&trade_owner&"',trade_addr='"&trade_addr&"',trade_uptae='"&trade_uptae&"',trade_upjong='"&trade_upjong&"',trade_tel='"&trade_tel&"',trade_fax='"&trade_fax&"',mg_group='"&mg_group&"',group_name='"&group_name&"',use_sw='"&use_sw&"',mod_id='"&user_id&"',mod_date=now() where trade_code ='"&trade_code&"'"
		dbconn.execute(sql)
	  else

		sql="select max(trade_code) as max_seq from trade"
		set rs=dbconn.execute(sql)
			
		if	isnull(rs("max_seq"))  then
			trade_code = "00001"
		  else
			max_seq = "0000" + cstr((int(rs("max_seq")) + 1))
			trade_code = right(max_seq,5)
		end if

		sql = "select * from trade where trade_no ='"&trade_no&"' or trade_name ='"&trade_name&"' or trade_full_name ='"&trade_full_name&"'"
		Set rs=DbConn.Execute(Sql)
		if rs.eof or rs.bof then
			sql="insert into trade (trade_code,trade_no,trade_name,bill_trade_code,bill_trade_name,trade_id,sales_type,trade_owner,trade_addr,trade_uptae,trade_upjong,trade_tel,trade_fax,mg_group,group_name,use_sw,reg_id,reg_date) values ('"&trade_code&"','"&trade_no&"','"&trade_name&"','"&bill_trade_code&"','"&bill_trade_name&"','"&trade_id&"','"&sales_type&"','"&trade_owner&"','"&trade_addr&"','"&trade_uptae&"','"&trade_upjong&"','"&trade_tel&"','"&trade_fax&"','"&mg_group&"','"&group_name&"','"&use_sw&"','"&user_id&"',now())"
			dbconn.execute(sql)
		  else
			response.write"<script language=javascript>"
			response.write"alert('이미 등록되어 있는 거래처입니다');"
			response.write"history.back();"
			response.write"</script>"
		end if
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "처리 되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"opener.document.frm.submit();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
