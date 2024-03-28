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

pmg_yymm="201505"
'view_condi = "케이원정보통신"
'view_condi = "에스유에이치"
'view_condi = "케이네트웍스"
'view_condi = "휴디스"
'view_condi = "코리아디엔씨"
pmg_id = "1" '1 2 4
pmg_date = "2015-05-31"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

i = 0
j = 0

Sql = "select * from pay_month_give where (pmg_yymm > '201301' )  and (pmg_id = '"+pmg_id+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"

'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    j = j + 1
	emp_no = rs("pmg_emp_no")
	
	pmg_yymm = rs("pmg_yymm")
	'pmg_id = rs("pmg_id")
	view_condi = rs("pmg_company")
	
	pmg_base_pay = rs("pmg_base_pay")
	pmg_meals_pay = rs("pmg_meals_pay")
	pmg_postage_pay = rs("pmg_postage_pay")
	pmg_re_pay = rs("pmg_re_pay")
	pmg_overtime_pay = rs("pmg_overtime_pay")
	pmg_car_pay = rs("pmg_car_pay")
	pmg_position_pay = rs("pmg_position_pay")
	pmg_custom_pay = rs("pmg_custom_pay")
	pmg_job_pay = rs("pmg_job_pay")
	pmg_job_support = rs("pmg_job_support")
	pmg_jisa_pay = rs("pmg_jisa_pay")
	pmg_long_pay = rs("pmg_long_pay")
	pmg_disabled_pay = rs("pmg_disabled_pay")
	
	pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay
				
	meals_pay = pmg_meals_pay
	car_pay = pmg_car_pay
	meals_tax_pay = 0
	car_tax_pay = 0
	if  meals_pay > 100000 then
	         meals_tax_pay = meals_pay - 100000
			 meals_pay =  100000
	end if
	if car_pay > 200000 then
	         car_tax_pay = car_pay - 200000
			 car_pay =  200000
	end if
	
	pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	pmg_tax_no = meals_pay + car_pay
	
		
	sql = "update pay_month_give set pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_give_total='"&pmg_give_total&"' where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
		
	dbconn.execute(sql)	 
		   

	    Rs.MoveNext()
  loop		
		response.write"<script language=javascript>"
		response.write"alert('급여에 과세/비과세 금액이 만들어 졌습니다..."&j&" - "&i&"');"		
		response.write"location.replace('insa_person_mg.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert(' 처리된 내역이없습니다...');"		
		response.write"location.replace('insa_person_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
