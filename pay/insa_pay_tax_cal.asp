<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim emp_name, rever_year, pmg_tax_yes, pmg_give_tot, incom_family_cnt
Dim in_pmg_id, err_rtn

emp_no = Request.Form("emp_no")
emp_name = Request.Form("pmg_emp_name")
rever_year = Request.Form("rever_year")
pmg_tax_yes = CLng(Request.Form("pmg_tax_yes"))
pmg_give_tot = CLng(Request.Form("pmg_give_tot"))
incom_family_cnt = CLng(Request.Form("incom_family_cnt"))
in_pmg_id = Request.Form("in_pmg_id")

err_rtn = ""

Dim rsNps, nps_emp, nps_com, nps_from, nps_to

'국민연금 요율
objBuilder.Append "SELECT emp_rate, com_rate, from_amt, to_amt "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5501' AND insu_class = '01';"

Set rsNps = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsNps.EOF Then
	nps_emp = FormatNumber(rsNps("emp_rate"), 3)
	nps_com = FormatNumber(rsNps("com_rate"), 3)
	nps_from = rsNps("from_amt")
	nps_to = rsNps("to_amt")
Else
	nps_emp = 0
	nps_com = 0
	nps_from = 0
	nps_to = 0
End If

rsNps.Close() : Set rsNps = Nothing

Dim rsNhis, nhis_emp, nhis_com, nhis_from, nhis_to
'건강보험 요율
'sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5502' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate, com_rate, from_amt, to_amt "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5502' AND insu_class = '01';"

Set rsNhis = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsNhis.EOF Then
	nhis_emp = FormatNumber(rsNhis("emp_rate"), 3)
	nhis_com = FormatNumber(rsNhis("com_rate"), 3)
	nhis_from = rsNhis("from_amt")
	nhis_to = rsNhis("to_amt")
Else
	nhis_emp = 0
	nhis_com = 0
	nhis_from = 0
	nhis_to = 0
End If
rsNhis.Close() : Set rsNhis = Nothing

Dim rsEpi, epi_emp, epi_com
'고용보험(실업) 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5503' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate, com_rate "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5503' AND insu_class = '01';"

Set rsEpi = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsEpi.EOF Then
	epi_emp = FormatNumber(rsEpi("emp_rate"), 3)
	epi_com = FormatNumber(rsEpi("com_rate"), 3)
Else
	epi_emp = 0
	epi_com = 0
End If
rsEpi.Close() : Set rsEpi = Nothing

Dim rsHap, long_hap
'장기요양보험 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5504' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5504' AND insu_class = '01';"

Set rsHap = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsHap.EOF Then
	long_hap = FormatNumber(rsHap("hap_rate"), 3)
Else
	long_hap = 0
End If
rsHap.Close() : Set rsHap = Nothing

'Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
'Set rs = DbConn.Execute(SQL)

'start_date = rs("emp_first_date")
'emp_in_date = rs("emp_in_date")
'rs.close()

Dim rsYear, pmg_base_pay, pmg_meals_pay, pmg_overtime_pay, incom_month_amount, incom_nps_amount
Dim incom_nhis_amount, incom_nps, incom_nhis, incom_wife_yn, incom_age20, incom_age60, incom_old
Dim incom_disab, incom_go_yn, incom_long_yn, inc_incom, rs_sod, de_income_tax, nps_amt, nhis_amt
Dim long_amt, epi_amt, we_tax, deduct_tot, de_nps_amt, de_nhis_amt, de_longcare_amt, de_epi_amt
Dim de_wetax, curr_amt, emp_in_date, start_date

'기본급/식대 가져오기
incom_family_cnt = 0

'Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
objBuilder.Append "SELECT emtt.emp_first_date, emtt.emp_in_date, "
objBuilder.Append "	pyit.incom_base_pay, pyit.incom_meals_pay, pyit.incom_overtime_pay, pyit.incom_month_amount, "
objBuilder.Append "	pyit.incom_family_cnt, pyit.incom_nps_amount, pyit.incom_nhis_amount, pyit.incom_nps, "
objBuilder.Append "	pyit.incom_nhis, pyit.incom_wife_yn, pyit.incom_age20, pyit.incom_age60, pyit.incom_old, "
objBuilder.Append "	pyit.incom_disab, pyit.incom_go_yn, incom_long_yn "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "LEFT OUTER JOIN pay_year_income AS pyit ON emtt.emp_no = pyit.incom_emp_no "
objBuilder.Append "	AND pyit.incom_year = '"&rever_year&"' "
objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"';"

Set rsYear = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsYear.EOF Then
	start_date = rsYear("emp_first_date")
	emp_in_date = rsYear("emp_in_date")

	pmg_base_pay = f_toString(rsYear("incom_base_pay"), 0)
	pmg_meals_pay = f_toString(rsYear("incom_meals_pay"), 0)
	pmg_overtime_pay = f_toString(rsYear("incom_overtime_pay"), 0)

	If f_toString(rsYear("incom_month_amount"), 0) = 0 Then
		incom_month_amount = rsYear("incom_base_pay") + rsYear("incom_overtime_pay")
	Else
		incom_month_amount = f_toString(rsYear("incom_month_amount"), 0)
	End If

	incom_family_cnt = f_toString(rsYear("incom_family_cnt"), 0)
	incom_nps_amount = f_toString(rsYear("incom_nps_amount"), 0)
	incom_nhis_amount = f_toString(rsYear("incom_nhis_amount"), 0)
	incom_nps = f_toString(rsYear("incom_nps"), 0)
	incom_nhis = f_toString(rsYear("incom_nhis"), 0)
	incom_wife_yn = f_toString(rsYear("incom_wife_yn"), 0)
	incom_age20 = f_toString(rsYear("incom_age20"), 0)
	incom_age60 = f_toString(rsYear("incom_age60"), 0)
	incom_old = f_toString(rsYear("incom_old"), 0)
	incom_disab = f_toString(rsYear("incom_disab"), 0)
	incom_go_yn = f_toString(rsYear("incom_go_yn"), "여")
	incom_long_yn = f_toString(rsYear("incom_long_yn"), "여")
Else
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
End If
rsYear.Close() : Set rsYear = Nothing

'if incom_family_cnt = 0 then
    incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + 1 + incom_age20 + incom_disab'본인포함 및 20세이하/장애인은 추가공제
'end if

inc_incom = 0

'If in_pmg_id = "1" Then
' 	Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
'Else
'	Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
'End If
objBuilder.Append "SELECT inc_st_amt, inc_incom1, inc_incom2, inc_incom3, inc_incom4, inc_incom5, "
objBuilder.Append "	inc_incom6, inc_incom7, inc_incom8, inc_incom9, inc_incom10, inc_incom11 "
objBuilder.Append "FROM pay_income_amount "
objBuilder.Append "WHERE ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) "
objBuilder.Append "	AND inc_yyyy = '"&rever_year&"';"

Set rs_sod = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs_sod.EOF Then
	inc_st_amt = Int(rs_sod("inc_st_amt"))

	If incom_family_cnt = 1 Then
	   inc_incom = rs_sod("inc_incom1")
	End If

	If incom_family_cnt = 2 Then
	   inc_incom = rs_sod("inc_incom2")
	End If

	If incom_family_cnt = 3 Then
	   inc_incom = rs_sod("inc_incom3")
	End If

	If incom_family_cnt = 4 Then
	   inc_incom = rs_sod("inc_incom4")
	End If

	If incom_family_cnt = 5 Then
	   inc_incom = rs_sod("inc_incom5")
	End If

	If incom_family_cnt = 6 Then
	   inc_incom = rs_sod("inc_incom6")
	End If

	If incom_family_cnt = 7 Then
	   inc_incom = rs_sod("inc_incom7")
	End If

	If incom_family_cnt = 8 Then
	   inc_incom = rs_sod("inc_incom8")
	End If

	If incom_family_cnt = 9 Then
	   inc_incom = rs_sod("inc_incom9")
	End If

	If incom_family_cnt = 10 Then
	   inc_incom = rs_sod("inc_incom10")
	End If

	If incom_family_cnt = 11 Then
	   inc_incom = rs_sod("inc_incom11")
	End If
End If
rs_sod.Close() : Set rs_sod = Nothing
DBConn.Close() : Set DBConn = Nothing

'소득세
de_income_tax = CInt(inc_incom)

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
If incom_long_yn = "여" Then
	long_amt = de_nhis_amt * (long_hap / 100)
	long_amt = CInt(long_amt)
	'long_amt = long_amt / 2
	de_longcare_amt = (CInt(long_amt / 10)) * 10
Else
	de_longcare_amt = 0
End If

'고용보험 계산 : 비과세 포함한 금액으로 계산
If incom_go_yn = "여" Then
	'epi_amt = inc_st_amt * (epi_emp / 100)
	epi_amt = pmg_tax_yes * (epi_emp / 100)
	epi_amt = CInt(epi_amt)
	de_epi_amt = (CInt(epi_amt / 10)) * 10
Else
	de_epi_amt = 0
End If



'지방소득세
we_tax = inc_incom * (10 / 100)
we_tax = CInt(we_tax)
de_wetax = (CInt(we_tax / 10)) * 10

deduct_tot = de_epi_amt + de_income_tax + de_wetax + de_nps_amt + de_nhis_amt + de_longcare_amt
curr_amt = pmg_tax_yes - deduct_tot

If err_rtn <> "" Then
	Response.Write err_rtn
Else
	Response.Write emp_in_date
	Response.Write "|" & in_pmg_id & "|" & de_epi_amt & "|" & de_income_tax & "|" & de_wetax & "|" & de_nps_amt & "|" & de_nhis_amt & "|" & de_longcare_amt & "|" & deduct_tot & "|" & curr_amt
	'response.write "|" & family_empno
End If
Response.End
%>