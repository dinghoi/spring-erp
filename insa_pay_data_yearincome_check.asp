<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

i = 0
j = 0
rever_year = "2015"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'국민연금 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5501' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nps_emp = formatnumber(rs_ins("emp_rate"),3)
		nps_com = formatnumber(rs_ins("com_rate"),3)
		nps_from = rs_ins("from_amt")
		nps_to = rs_ins("to_amt")
   else
		nps_emp = 0
		nps_com = 0
		nps_from = 0
		nps_to = 0
end if
rs_ins.close()

'건강보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5502' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nhis_emp = formatnumber(rs_ins("emp_rate"),3)
		nhis_com = formatnumber(rs_ins("com_rate"),3)
		nhis_from = rs_ins("from_amt")
		nhis_to = rs_ins("to_amt")
   else
		nhis_emp = 0  
		nhis_com = 0
		nhis_from = 0
		his_to = 0
end if
rs_ins.close()

'고용보험(실업) 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5503' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	epi_emp = formatnumber(rs_ins("emp_rate"),3)
		epi_com = formatnumber(rs_ins("com_rate"),3)
   else
		epi_emp = 0  
		epi_com = 0
end if
rs_ins.close()

'장기요양보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5504' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	long_hap = formatnumber(rs_ins("hap_rate"),3)
   else
		long_hap = 0  
end if
rs_ins.close()

Sql = "select * from pay_year_income where incom_year = '"&rever_year&"'"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    incom_emp_no = rs("incom_emp_no")
	incom_year = rs("incom_year")
	
 '건강보험 계산
    incom_nhis_amount = int(rs("incom_nhis_amount"))
    nhis_amt = incom_nhis_amount * (nhis_emp / 100)
    nhis_amt = int(nhis_amt)
    incom_nhis = (int(nhis_amt / 10)) * 10

    sql = "update pay_year_income set incom_nhis='"&incom_nhis&"' where (incom_emp_no = '"&incom_emp_no&"') and (incom_year = '"&incom_year&"')"
		
	dbconn.execute(sql)	 

    j = j + 1 

	
	Rs.MoveNext()
  loop		
		response.write"<script language=javascript>"
		response.write"alert('건강보험 월납부금액이 변경 되었습니다....(갱신-"&j&").');"		
		response.write"location.replace('insa_person_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
