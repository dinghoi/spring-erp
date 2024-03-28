<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'on error resume next
response.charset = "euc-kr"

Dim cnt
emp_no = request.form("emp_no")
emp_name = request.form("pmg_emp_name")
rever_year = request.form("rever_year")
pmg_tax_yes = int(request.form("pmg_tax_yes"))
pmg_give_tot = int(request.form("pmg_give_tot"))
incom_family_cnt = int(request.form("incom_family_cnt"))
in_pmg_id = request.form("in_pmg_id")

err_rtn = ""

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
Set rs = DbConn.Execute(SQL)

start_date = rs("emp_first_date")
emp_in_date = rs("emp_in_date")
rs.close()

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

'기본급/식대 가져오기
incom_family_cnt = 0
Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
Set Rs_year = DbConn.Execute(SQL)
if not Rs_year.eof then
    	pmg_base_pay = Rs_year("incom_base_pay")
		pmg_meals_pay = Rs_year("incom_meals_pay")
		pmg_overtime_pay = Rs_year("incom_overtime_pay")
		if Rs_year("incom_month_amount") = 0 or isnull(Rs_year("incom_month_amount")) then
		        incom_month_amount = Rs_year("incom_base_pay") + Rs_year("incom_overtime_pay")
		   else
		        incom_month_amount = Rs_year("incom_month_amount")
		end if
		incom_family_cnt = Rs_year("incom_family_cnt")
		incom_nps_amount = Rs_year("incom_nps_amount")
		incom_nhis_amount = Rs_year("incom_nhis_amount")
		incom_nps = Rs_year("incom_nps")
		incom_nhis = Rs_year("incom_nhis")
		incom_wife_yn = int(Rs_year("incom_wife_yn"))
		incom_age20 = Rs_year("incom_age20")
		incom_age60 = Rs_year("incom_age60")
		incom_old = Rs_year("incom_old")
		incom_disab = Rs_year("incom_disab")
		incom_go_yn = Rs_year("incom_go_yn")
		incom_long_yn = Rs_year("incom_long_yn")
   else
		pmg_base_pay = 0
		pmg_meals_pay = 0
		pmg_overtime_pay = 0
		incom_month_amount = 0
		incom_family_cnt = 0
		incom_nps_amount = 0
		incom_nhis_amount = 0
		incom_nps = 0
		incom_nhis = 0
		incom_wife_yn = 0
		incom_age20 = 0
		incom_age60 = 0
		incom_old = 0
		incom_disab = 0
		incom_go_yn = "여"
		incom_long_yn = "여"
end if
Rs_year.close()

'if incom_family_cnt = 0 then
    incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + 1 + incom_age20 + incom_disab'본인포함 및 20세이하/장애인은 추가공제
'end if

inc_incom = 0
if in_pmg_id = "1" then
'     Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' BETWEEN inc_from_amt and inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
	 Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
   else
'     Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' BETWEEN inc_from_amt and inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
	 Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
end if
Set Rs_sod = DbConn.Execute(SQL)
if not Rs_sod.eof then
	inc_st_amt = int(Rs_sod("inc_st_amt"))
	if incom_family_cnt = 1 then
	   inc_incom = Rs_sod("inc_incom1")
	end if
	if incom_family_cnt = 2 then
	   inc_incom = Rs_sod("inc_incom2")
	end if
	if incom_family_cnt = 3 then
	   inc_incom = Rs_sod("inc_incom3")
	end if
	if incom_family_cnt = 4 then
	   inc_incom = Rs_sod("inc_incom4")
	end if
	if incom_family_cnt = 5 then
	   inc_incom = Rs_sod("inc_incom5")
	end if
	if incom_family_cnt = 6 then
	   inc_incom = Rs_sod("inc_incom6")
	end if
	if incom_family_cnt = 7 then
	   inc_incom = Rs_sod("inc_incom7")
	end if
	if incom_family_cnt = 8 then
	   inc_incom = Rs_sod("inc_incom8")
	end if
	if incom_family_cnt = 9 then
	   inc_incom = Rs_sod("inc_incom9")
	end if
	if incom_family_cnt = 10 then
	   inc_incom = Rs_sod("inc_incom10")
	end if
	if incom_family_cnt = 11 then
	   inc_incom = Rs_sod("inc_incom11")
	end if
end if
Rs_sod.close()

'소득세
de_income_tax = int(inc_incom)

nps_amt = 0 '국민연금
nhis_amt = 0 '건강보험
long_amt = 0 '장기요양보험
epi_amt = 0 '고용보험
we_tax = 0 '지방소득세
deduct_tot = 0

'국민연금 계산
'nps_amt = incom_nps_amount * (nps_emp / 100)
'nps_amt = int(nps_amt)
'de_nps_amt = (int(nps_amt / 10)) * 10
de_nps_amt = incom_nps

'건강보험 계산
'nhis_amt = incom_nhis_amount * (nhis_emp / 100)
'nhis_amt = int(nhis_amt)
'de_nhis_amt = (int(nhis_amt / 10)) * 10
de_nhis_amt = incom_nhis

'장기요양보험 계산
if incom_long_yn = "여" then
        long_amt = de_nhis_amt * (long_hap / 100)
        long_amt = Int(long_amt)
        'long_amt = long_amt / 2
        de_longcare_amt = (Int(long_amt / 10)) * 10
   else
        de_longcare_amt = 0
end if

'고용보험 계산 : 비과세 포함한 금액으로 계산
if incom_go_yn = "여" then
        'epi_amt = inc_st_amt * (epi_emp / 100)
		epi_amt = pmg_tax_yes * (epi_emp / 100)
        epi_amt = int(epi_amt)
        de_epi_amt = (int(epi_amt / 10)) * 10
   else
		de_epi_amt = 0
end if

'지방소득세
we_tax = inc_incom * (10 / 100)
we_tax = int(we_tax)
de_wetax = (int(we_tax / 10)) * 10

deduct_tot = de_epi_amt + de_income_tax + de_wetax + de_nps_amt + de_nhis_amt + de_longcare_amt
curr_amt = pmg_tax_yes - deduct_tot

if err_rtn <> "" then
	response.write err_rtn
else
	response.write emp_in_date
	response.write "|" & in_pmg_id & "|" & de_epi_amt & "|" & de_income_tax & "|" & de_wetax & "|" & de_nps_amt & "|" & de_nhis_amt & "|" & de_longcare_amt & "|" & deduct_tot & "|" & curr_amt
	'response.write "|" & family_empno
end if

%>