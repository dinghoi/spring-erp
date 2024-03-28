<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	d_emp_no = request.form("d_emp_no")
	d_year = request.form("d_year")
	d_seq = request.form("d_seq")
	
	d_emp_name = request.form("d_emp_name")
	d_name = request.form("d_name")
	d_rel = request.form("d_rel")
	d_person_no = request.form("d_person_no")
	d_data_gubun = request.form("d_data_gubun")
	d_trade_no = request.form("d_trade_no")
	d_trade_name = request.form("d_trade_name")
	d_nts_chk = request.form("d_nts_chk")

	d_cnt =int(request.form("d_cnt"))
	d_amt =int(request.form("d_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_donation set d_rel='"&d_rel&"',d_name='"&d_name&"',d_trade_no='"&d_trade_no&"',d_trade_name='"&d_trade_name&"',d_nts_chk='"&d_nts_chk&"',d_data_gubun='"&d_data_gubun&"',d_cnt='"&d_cnt&"',d_amt='"&d_amt&"' where d_year ='"&d_year&"' and d_emp_no = '"&d_emp_no&"' and d_person_no = '"&d_person_no&"' and d_seq = '"&d_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(d_seq) as max_seq from pay_yeartax_donation where d_year='" + d_year + "' and d_emp_no='" + d_emp_no + "' and d_person_no='" + d_person_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			d_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			d_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_donation (d_year,d_emp_no,d_person_no,d_seq,d_rel,d_name,d_data_gubun,d_trade_no,d_trade_name,d_cnt,d_amt,d_nts_chk) values "
		sql = sql +	" ('"&d_year&"','"&d_emp_no&"','"&d_person_no&"','"&d_seq&"','"&d_rel&"','"&d_name&"','"&d_data_gubun&"','"&d_trade_no&"','"&d_trade_name&"','"&d_cnt&"','"&d_amt&"','"&d_nts_chk&"')"
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
