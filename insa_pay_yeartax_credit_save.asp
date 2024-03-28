<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	c_id = request.form("c_id")
	
	c_emp_no = request.form("c_emp_no")
	c_year = request.form("c_year")
	c_seq = request.form("c_seq")
	
	c_emp_name = request.form("c_emp_name")
	cc_name = request.form("cc_name")
	c_rel = request.form("c_rel")
	c_person_no = request.form("c_person_no")
	c_market = request.form("c_market")
	c_transit = request.form("c_transit")

	c_nts_amt =int(request.form("c_nts_amt"))
	c_other_amt =int(request.form("c_other_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_credit set c_rel='"&c_rel&"',cc_name='"&cc_name&"',c_market='"&c_market&"',c_transit='"&c_transit&"',c_nts_amt='"&c_nts_amt&"',c_other_amt='"&c_other_amt&"' where c_year ='"&c_year&"' and c_emp_no = '"&c_emp_no&"' and c_person_no = '"&c_person_no&"' and c_id = '"&c_id&"' and c_seq = '"&c_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(c_seq) as max_seq from pay_yeartax_credit where c_year='" + c_year + "' and c_emp_no='" + c_emp_no + "' and c_person_no='" + c_person_no + "' and c_id='" + c_id + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			c_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			c_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_credit (c_year,c_emp_no,c_person_no,c_id,c_seq,c_rel,cc_name,c_market,c_transit,c_nts_amt,c_other_amt) values "
		sql = sql +	" ('"&c_year&"','"&c_emp_no&"','"&c_person_no&"','"&c_id&"','"&c_seq&"','"&c_rel&"','"&cc_name&"','"&c_market&"','"&c_transit&"','"&c_nts_amt&"','"&c_other_amt&"')"
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
