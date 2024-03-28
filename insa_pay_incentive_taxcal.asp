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
incom_family_cnt = int(request.form("incom_family_cnt"))
incom_go_yn = request.form("incom_go_yn")
in_pmg_id = request.form("in_pmg_id")

err_rtn = ""

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
'rever_year = mid(cstr(pmg_yymm),1,4)
'incom_family_cnt = 1

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


'기본급/식대 가져오기
incom_family_cnt = 0
Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
Set Rs_year = DbConn.Execute(SQL)
if not Rs_year.eof then
		if Rs_year("incom_month_amount") = 0 or isnull(Rs_year("incom_month_amount")) then
		        incom_month_amount = Rs_year("incom_base_pay") + Rs_year("incom_overtime_pay")
		   else
		        incom_month_amount = Rs_year("incom_month_amount")
		end if
		incom_family_cnt = Rs_year("incom_family_cnt")
		incom_wife_yn = int(Rs_year("incom_wife_yn"))
		incom_age20 = Rs_year("incom_age20")
		incom_age60 = Rs_year("incom_age60")
		incom_old = Rs_year("incom_old")
		incom_disab = Rs_year("incom_disab")
		incom_go_yn = Rs_year("incom_go_yn")
   else
		incom_month_amount = 0
		incom_family_cnt = 0
		incom_wife_yn = 0
		incom_age20 = 0
		incom_age60 = 0
		incom_old = 0
		incom_disab = 0
		incom_go_yn = "여"
end if
Rs_year.close()

'if incom_family_cnt = 0 then
    incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + 1 + incom_age20 + incom_disab'본인포함 및 20세이하/장애인은 추가공제
'end if

inc_incom = 0
Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
Set Rs_sod = DbConn.Execute(SQL)
if not Rs_sod.eof then
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
inc_incom = int(inc_incom)

de_epi_amt = 0
we_tax = 0
deduct_tot = 0


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
we_tax = (int(we_tax / 10)) * 10 

if in_pmg_id = "4" then  ' 연차수당인경우 소득세/주민세 공제 안함
    inc_incom = 0
	we_tax = 0
end if

deduct_tot = de_epi_amt + inc_incom + we_tax
curr_amt = pmg_tax_yes - deduct_tot

if err_rtn <> "" then
	response.write err_rtn
else
	response.write emp_in_date
	response.write "|" & rever_year & "|" & de_epi_amt & "|" & inc_incom & "|" & we_tax & "|" & deduct_tot & "|" & curr_amt 
	'response.write "|" & family_empno 
end if

%>