<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	m_emp_no = request.form("m_emp_no")
	m_year = request.form("m_year")
	m_seq = request.form("m_seq")
	
	m_emp_name = request.form("m_emp_name")
	m_name = request.form("m_name")
	m_rel = request.form("m_rel")
	m_nation = request.form("m_nation")
	m_pensioner = request.form("m_pensioner")
	m_witak = request.form("m_witak")
	m_person_no = request.form("m_person_no")
	m_data_gubun = request.form("m_data_gubun")
	m_disab = request.form("m_disab")
	m_age65 = request.form("m_age65")
	m_trade_no = request.form("m_trade_no")
	m_trade_name = request.form("m_trade_name")
	m_eye = request.form("m_eye")

	m_cnt =int(request.form("m_cnt"))
	m_amt =int(request.form("m_amt"))
	
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_yeartax_medical set m_rel='"&m_rel&"',m_name='"&m_name&"',m_nation='"&m_nation&"',m_pensioner='"&m_pensioner&"',m_witak='"&m_witak&"',m_disab='"&m_disab&"',m_age65='"&m_age65&"',m_trade_no='"&m_trade_no&"',m_trade_name='"&m_trade_name&"',m_eye='"&m_eye&"',m_data_gubun='"&m_data_gubun&"',m_cnt='"&m_cnt&"',m_amt='"&m_amt&"' where m_year ='"&m_year&"' and m_emp_no = '"&m_emp_no&"' and m_person_no = '"&m_person_no&"' and m_seq = '"&m_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(m_seq) as max_seq from pay_yeartax_medical where m_year='" + m_year + "' and m_emp_no='" + m_emp_no + "' and m_person_no='" + m_person_no + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			m_seq = "01"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			m_seq = right(max_seq,2)
		end if

		sql = "insert into pay_yeartax_medical (m_year,m_emp_no,m_person_no,m_seq,m_rel,m_name,m_nation,m_pensioner,m_witak,m_disab,m_age65,m_trade_no,m_trade_name,m_eye,m_data_gubun,m_cnt,m_amt) values "
		sql = sql +	" ('"&m_year&"','"&m_emp_no&"','"&m_person_no&"','"&m_seq&"','"&m_rel&"','"&m_name&"','"&m_nation&"','"&m_pensioner&"','"&m_witak&"','"&m_disab&"','"&m_age65&"','"&m_trade_no&"','"&m_trade_name&"','"&m_eye&"','"&m_data_gubun&"','"&m_cnt&"','"&m_amt&"')"
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
