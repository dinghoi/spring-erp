<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	hm_emp_no = request.form("hm_emp_no")
	hm_year = request.form("hm_year")
	hm_seq = request.form("hm_seq")
	
	hm_emp_name = request.form("hm_emp_name")
	hm_from_date = request.form("hm_from_date")
	hm_to_date = request.form("hm_to_date")

	hm_month_amt =int(request.form("hm_month_amt"))
	
	hm_data_gubun = request.form("hm_data_gubun")
	hm_trade_name = request.form("hm_trade_name")
	hm_trade_no = request.form("hm_trade_no")
	hm_house_type = request.form("hm_house_type")
	hm_size = request.form("hm_size")
	hm_addr = request.form("hm_addr")
	hm_lender = request.form("hm_lender")
	hm_lender_person = request.form("hm_lender_person")
	hm_lender_from = request.form("hm_lender_from")
	hm_lender_to = request.form("hm_lender_to")
	if hm_lender_from = "" or isnull(hm_lender_from) then
	   hm_lender_from = "1900-01-01"
	end if
	if hm_lender_to = "" or isnull(hm_lender_to) then
	   hm_lender_to = "1900-01-01"
	end if
	hm_lender_rate = request.form("hm_lender_rate")
	hm_lender_amt =int(request.form("hm_lender_amt"))
	hm_lender_rate_amt =int(request.form("hm_lender_rate_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_house_m set hm_data_gubun='"&hm_data_gubun&"',hm_trade_name='"&hm_trade_name&"',hm_trade_no='"&hm_trade_no&"',hm_house_type='"&hm_house_type&"',hm_size='"&hm_size&"',hm_addr='"&hm_addr&"',hm_lender='"&hm_lender&"',hm_lender_person='"&hm_lender_person&"',hm_lender_from='"&hm_lender_from&"',hm_lender_to='"&hm_lender_to&"',hm_lender_rate='"&hm_lender_rate&"',hm_lender_amt='"&hm_lender_amt&"',hm_lender_rate_amt='"&hm_lender_rate_amt&"',hm_from_date='"&hm_from_date&"',hm_to_date='"&hm_to_date&"',hm_month_amt='"&hm_month_amt&"' where hm_year ='"&hm_year&"' and hm_emp_no = '"&hm_emp_no&"' and hm_seq = '"&hm_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(hm_seq) as max_seq from pay_yeartax_house_m where hm_year='" + hm_year + "' and hm_emp_no='" + hm_emp_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			hm_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			hm_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_house_m (hm_year,hm_emp_no,hm_seq,hm_emp_name,hm_data_gubun,hm_trade_name,hm_trade_no,hm_house_type,hm_size,hm_addr,hm_from_date,hm_to_date,hm_month_amt,hm_lender,hm_lender_person,hm_lender_from,hm_lender_to,hm_lender_rate,hm_lender_amt,hm_lender_rate_amt) values "
		sql = sql +	" ('"&hm_year&"','"&hm_emp_no&"','"&hm_seq&"','"&hm_emp_name&"','"&hm_data_gubun&"','"&hm_trade_name&"','"&hm_trade_no&"','"&hm_house_type&"','"&hm_size&"','"&hm_addr&"','"&hm_from_date&"','"&hm_to_date&"','"&hm_month_amt&"','"&hm_lender&"','"&hm_lender_person&"','"&hm_lender_from&"','"&hm_lender_to&"','"&hm_lender_rate&"','"&hm_lender_amt&"','"&hm_lender_rate_amt&"')"
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
