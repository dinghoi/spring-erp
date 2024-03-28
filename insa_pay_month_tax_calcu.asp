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

pmg_yymm=Request.form("pmg_yymm1")
view_condi=Request.form("view_condi1")
in_empno=Request.form("in_empno1")
in_name=Request.form("in_name1")
owner_view=Request.form("owner_view1")

'response.write(pmg_yymm)
'response.write(view_condi)
'response.write(in_empno)
'response.write(in_name)
'response.write(owner_view)
'response.End

rever_year = mid(cstr(pmg_yymm),1,4)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
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

if owner_view = "T" then 
       Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_no = '"+in_empno+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
   else
       Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
end if
Rs.Open Sql, Dbconn, 1

pmg_emp_no = rs("pmg_emp_no")

do until rs.eof
	pmg_emp_no = rs("pmg_emp_no")
	emp_no = rs("pmg_emp_no")
    pmg_company = rs("pmg_company")
	pmg_date = rs("pmg_date")
	pmg_emp_name = rs("pmg_emp_name")
	pmg_org_code = rs("pmg_org_code")
	pmg_org_name = rs("pmg_org_name")
	pmg_emp_type = rs("pmg_emp_type")
	pmg_grade = rs("pmg_grade")
	pmg_position = rs("pmg_position")	
	
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
	'pmg_family_pay = rs("pmg_family_pay")
	'pmg_school_pay = rs("pmg_school_pay")
	'pmg_qual_pay = rs("pmg_qual_pay")
	'pmg_other_pay1 = rs("pmg_other_pay1")
	'pmg_other_pay2 = rs("pmg_other_pay2")
	'pmg_other_pay3 = rs("pmg_other_pay3")
	pmg_tax_yes = rs("pmg_tax_yes")
	pmg_tax_no = rs("pmg_tax_no")
	pmg_tax_reduced = rs("pmg_tax_reduced")
	pmg_give_total = rs("pmg_give_total")	

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

    incom_family_cnt = 0

    Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&pmg_emp_no&"' and incom_year = '"&rever_year&"'"
    Set Rs_year = DbConn.Execute(SQL)
    if not Rs_year.eof then
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
         Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
       else
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

	Sql = "SELECT * FROM emp_master where emp_no = '"&pmg_emp_no&"'"
    Set rs_emp = DbConn.Execute(SQL)
    if not rs_emp.eof then
	   emp_end_date = rs_emp("emp_end_date")
	   if emp_end_date = "" or isnull(emp_end_date) or emp_end_date = "1900-01-01" then
        	emp_first_date = rs_emp("emp_first_date")
		    emp_in_date = rs_emp("emp_in_date")
		    pmg_emp_type = rs_emp("emp_type")
		    pmg_grade = rs_emp("emp_grade")
		    pmg_position = rs_emp("emp_position")
		    pmg_company = rs_emp("emp_company")
		    pmg_bonbu = rs_emp("emp_bonbu")
		    pmg_saupbu = rs_emp("emp_saupbu")
		    pmg_team = rs_emp("emp_team")
		    pmg_org_code = rs_emp("emp_org_code")
		    pmg_org_name = rs_emp("emp_org_name")
		    pmg_reside_place = rs_emp("emp_reside_place")
		    pmg_reside_company = rs_emp("emp_reside_company")
			cost_center = rs_emp("cost_center")	  
		    cost_group = rs_emp("cost_group")
        end if
	end if
    rs_emp.close()  
	
    Sql = "SELECT * FROM pay_bank_account where emp_no = '"&pmg_emp_no&"'"
    Set rs_bnk = DbConn.Execute(SQL)
    if not rs_bnk.eof then
           pmg_bank_name = rs_bnk("bank_name")
           pmg_account_no = rs_bnk("account_no")
		   pmg_account_holder = rs_bnk("account_holder")
	   else
           pmg_bank_name = ""
		   pmg_account_no = ""
		   pmg_account_holder = ""
    end if
    rs_bnk.close() 

    sql = "Update pay_month_give set pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_give_total='"&pmg_give_total&"',pmg_bank_name='"&pmg_bank_name&"',pmg_account_no='"&pmg_account_no&"',pmg_account_holder='"&pmg_account_holder&"',pmg_mod_user='"&emp_user&"',pmg_mod_date=now() where pmg_yymm = '"&pmg_yymm&"' and pmg_id = '1' and pmg_emp_no = '"&pmg_emp_no&"' and pmg_company = '"&pmg_company&"'"
		dbconn.execute(sql)

	
	Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+pmg_emp_no+"') and (de_company = '"+pmg_company+"')"
    Set Rs_dct = DbConn.Execute(SQL)
	if not Rs_dct.eof then	
           de_other_amt1 = Rs_dct("de_other_amt1")
           de_sawo_amt = Rs_dct("de_sawo_amt")
           de_hyubjo_amt = Rs_dct("de_hyubjo_amt")
           de_school_amt = Rs_dct("de_school_amt")
           de_nhis_bla_amt = Rs_dct("de_nhis_bla_amt")
           de_long_bla_amt = Rs_dct("de_long_bla_amt")
		   de_year_incom_tax = Rs_dct("de_year_incom_tax")
           de_year_wetax = Rs_dct("de_year_wetax")	
		   de_year_incom_tax2 = Rs_dct("de_year_incom_tax2")
           de_year_wetax2 = Rs_dct("de_year_wetax2")	
		   
		   de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax + de_other_amt1 + de_sawo_amt + de_hyubjo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2
		   
		   sql = "Update pay_month_deduct set de_nps_amt='"&de_nps_amt&"',de_nhis_amt ='"&de_nhis_amt&"',de_epi_amt ='"&de_epi_amt&"',de_longcare_amt ='"&de_longcare_amt&"',de_income_tax='"&de_income_tax&"',de_wetax='"&de_wetax&"',de_deduct_total='"&de_deduct_total&"',de_mod_user='"&emp_user&"',de_mod_date=now() where de_yymm = '"&pmg_yymm&"' and de_id = '1' and de_emp_no = '"&pmg_emp_no&"' and de_company = '"&pmg_company&"'"
		dbconn.execute(sql)
		   
	   else
	       de_other_amt1 = 0
		   de_special_tax = 0
           de_saving_amt = 0
           de_sawo_amt = 0
           de_johab_amt = 0
           de_hyubjo_amt = 0
           de_school_amt = 0
           de_nhis_bla_amt = 0
           de_long_bla_amt = 0
		   de_year_incom_tax = 0
           de_year_wetax = 0
		   de_year_incom_tax2 = 0
           de_year_wetax2 = 0
		   
		   de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax + de_other_amt1 + de_sawo_amt + de_hyubjo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2
		   
		   sql="insert into pay_month_deduct (de_yymm,de_id,de_emp_no,de_company,de_date,de_emp_name,de_emp_type,de_org_code,de_org_name,de_bonbu,de_saupbu,de_team,de_reside_place,de_reside_company,de_grade,de_position,de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_income_tax,de_year_wetax,de_year_income_tax2,de_year_wetax2,de_other_amt1,de_saving_amt,de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total,cost_group,cost_center,de_reg_date,de_reg_user) values ('"&pmg_yymm&"','1','"&pmg_emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&de_nps_amt&"','"&de_nhis_amt&"','"&de_epi_amt&"','"&de_longcare_amt&"','"&de_income_tax&"','"&de_wetax&"','"&de_year_income_tax&"','"&de_year_wetax&"','"&de_year_income_tax2&"','"&de_year_wetax2&"','"&de_other_amt1&"','"&de_saving_amt&"','"&de_sawo_amt&"','"&de_johab_amt&"','"&de_hyubjo_amt&"','"&de_school_amt&"','"&de_nhis_bla_amt&"','"&de_long_bla_amt&"','"&de_deduct_total&"','"&cost_group&"','"&cost_center&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
	   
    end if 
		
	Rs.MoveNext()
loop

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		'dbconn.CommitTrans
		end_msg = sms_msg + "급여 세금계산 데이타가 만들어 졌습니다..." 
	end if
 
    url = "insa_pay_month_pay_mg.asp?ck_sw=y&view_condi=" + view_condi + "&pmg_yymm="+ pmg_yymm


	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"location.replace('insa_pay_month_pay_mg.asp');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

dbconn.Close()
Set dbconn = Nothing
	
%>
