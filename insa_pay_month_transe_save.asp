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
pmg_yymm_to=Request.form("pmg_yymm_to1")
pmg_date=Request.form("to_date1")

'response.write(pmg_yymm)
'response.write(view_condi)
'response.write(pmg_yymm_to)
'response.write(to_date)
'response.End

'당월 입사/퇴사일이 15일 이전이면 당월 급여대상임
st_es_date = mid(cstr(pmg_yymm_to),1,4) + "-" + mid(cstr(pmg_yymm_to),5,2) + "-" + "01"
st_in_date = mid(cstr(pmg_yymm_to),1,4) + "-" + mid(cstr(pmg_yymm_to),5,2) + "-" + "16"
rever_year = mid(cstr(pmg_yymm_to),1,4) '귀속년도

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

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

' 급여지급월의 15일까지 입사자 당월급여처리를 위한 급여데이타 생성(전월급여지급이 없음)	
if view_condi = "전체" then
		   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
       else	   
           Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"')  and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
end if
Rs.Open Sql, Dbconn, 1

if not Rs.eof then
   do until Rs.eof
          emp_no = rs("emp_no")
		  emp_company = rs("emp_company")
		  emp_name = rs("emp_name")
		  pmg_emp_no = rs("emp_no")
	      pmg_emp_name = rs("emp_name")
		  pmg_in_date = rs("emp_in_date")
		  pmg_emp_type = rs("emp_type")
		  pmg_grade = rs("emp_grade")
		  pmg_position = rs("emp_position")
		  pmg_company = rs("emp_company")
		  pmg_bonbu = rs("emp_bonbu")
		  pmg_saupbu = rs("emp_saupbu")
		  pmg_team = rs("emp_team")
		  pmg_org_code = rs("emp_org_code")
		  pmg_org_name = rs("emp_org_name")
		  pmg_reside_place = rs("emp_reside_place")
		  pmg_reside_company = rs("emp_reside_company")	
		  cost_group = rs("cost_group")
		  cost_center = rs("cost_center")

          sql = "select * from pay_month_give where (pmg_yymm = '"&pmg_yymm&"' ) and (pmg_id = '1') and (pmg_emp_no = '"&emp_no&"') and (pmg_company = '"&emp_company&"')"
		  Set Rs_give = DbConn.Execute(SQL)
	      if not Rs_give.eof then	
		         pmg_company = Rs_give("pmg_company")
		         pmg_bonbu = Rs_give("pmg_bonbu")
		         pmg_saupbu = Rs_give("pmg_saupbu")
		         pmg_team = Rs_give("pmg_team")
		         pmg_org_name = Rs_give("pmg_org_name")
				 
		         pmg_base_pay = Rs_give("pmg_base_pay")
	             pmg_meals_pay = Rs_give("pmg_meals_pay")
	             pmg_postage_pay = Rs_give("pmg_postage_pay")
	             pmg_re_pay = Rs_give("pmg_re_pay")
	             pmg_overtime_pay = Rs_give("pmg_overtime_pay")
	             pmg_car_pay = Rs_give("pmg_car_pay")
	             pmg_position_pay = Rs_give("pmg_position_pay")
	             pmg_custom_pay = Rs_give("pmg_custom_pay")
	             pmg_job_pay = Rs_give("pmg_job_pay")
	             pmg_job_support = Rs_give("pmg_job_support")
	             pmg_jisa_pay = Rs_give("pmg_jisa_pay")
	             pmg_long_pay = Rs_give("pmg_long_pay")
	             pmg_disabled_pay = Rs_give("pmg_disabled_pay")
	             pmg_family_pay = Rs_give("pmg_family_pay")
	             pmg_school_pay = Rs_give("pmg_school_pay")
	             pmg_qual_pay = Rs_give("pmg_qual_pay")
	             pmg_other_pay1 = Rs_give("pmg_other_pay1")
	             pmg_other_pay2 = Rs_give("pmg_other_pay2")
	             pmg_other_pay3 = Rs_give("pmg_other_pay3")
	             pmg_tax_yes = Rs_give("pmg_tax_yes")
	             pmg_tax_no = Rs_give("pmg_tax_no")
	             pmg_tax_reduced = Rs_give("pmg_tax_reduced")
	             pmg_give_total = Rs_give("pmg_give_total")	
				 
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
				 
				 Sql = "select * from pay_month_deduct where (de_yymm = '"&pmg_yymm&"' ) and (de_id = '1') and (de_emp_no = '"&emp_no&"') and (de_company = '"&pmg_company&"')"
                 Set Rs_dct = DbConn.Execute(SQL)
		         if not Rs_dct.eof then
				        de_nps_amt = int(Rs_dct("de_nps_amt"))
                        de_nhis_amt = int(Rs_dct("de_nhis_amt"))
                        de_epi_amt = int(Rs_dct("de_epi_amt"))
                        de_longcare_amt = int(Rs_dct("de_longcare_amt"))
                        de_income_tax = int(Rs_dct("de_income_tax"))
                        de_wetax = int(Rs_dct("de_wetax"))
				        de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                        de_year_wetax = int(Rs_dct("de_year_wetax"))
						de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
                        de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
                        de_other_amt1 = int(Rs_dct("de_other_amt1"))
						de_special_tax = Rs_dct("de_special_tax")
                        de_saving_amt = Rs_dct("de_saving_amt")
                        de_sawo_amt = int(Rs_dct("de_sawo_amt"))
						de_johab_amt = Rs_dct("de_johab_amt")
                        de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                        de_school_amt = int(Rs_dct("de_school_amt"))
                        de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                        de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
                        de_deduct_total = int(Rs_dct("de_deduct_total"))	
                     else
				        de_nps_amt = 0
                        de_nhis_amt = 0
                        de_epi_amt = 0
                        de_longcare_amt = 0
                        de_income_tax = 0
                        de_wetax = 0
				        de_year_incom_tax = 0
				        de_year_wetax = 0
						de_year_incom_tax2 = 0
				        de_year_wetax2 = 0
                        de_other_amt1 = 0
                        de_special_tax = 0
                        de_saving_amt = 0
                        de_sawo_amt = 0
						de_johab_amt = 0
                        de_hyubjo_amt = 0
                        de_school_amt = 0
                        de_nhis_bla_amt = 0
                        de_long_bla_amt = 0
                        de_deduct_total = 0
                 end if
                 Rs_dct.close()

             else
                 pmg_base_pay = 0
	             pmg_meals_pay = 0
	             pmg_postage_pay = 0
	             pmg_re_pay = 0
	             pmg_overtime_pay = 0
	             pmg_car_pay = 0
	             pmg_position_pay = 0
	             pmg_custom_pay = 0
	             pmg_job_pay = 0
	             pmg_job_support = 0
	             pmg_jisa_pay = 0
	             pmg_long_pay = 0
	             pmg_disabled_pay = 0
	             pmg_family_pay = 0
	             pmg_school_pay = 0
	             pmg_qual_pay = 0
	             pmg_other_pay1 = 0
	             pmg_other_pay2 = 0
	             pmg_other_pay3 = 0
	             pmg_tax_yes = 0
	             pmg_tax_no = 0
	             pmg_tax_reduced = 0
	             pmg_give_total = 0
			
			     de_nps_amt = 0
                 de_nhis_amt = 0
                 de_epi_amt = 0
                 de_longcare_amt = 0
                 de_income_tax = 0
                 de_wetax = 0
			     de_year_incom_tax = 0
			     de_year_wetax = 0
				 de_year_incom_tax2 = 0
			     de_year_wetax2 = 0
                 de_other_amt1 = 0
                 de_special_tax = 0
                 de_saving_amt = 0
                 de_sawo_amt = 0
				 de_johab_amt = 0
                 de_hyubjo_amt = 0
                 de_school_amt = 0
                 de_nhis_bla_amt = 0
                 de_long_bla_amt = 0
                 de_deduct_total = 0
				 
				 '기본급/식대등 가져오기
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
		               incom_go_yn = "여"
		               incom_long_yn = "여"
		               incom_wife_yn = 0
		               incom_age20 = 0
		               incom_age60 = 0
		               incom_old = 0
                 end if
                 Rs_year.close()

                 pmg_tax_yes = pmg_base_pay + pmg_overtime_pay
                 pmg_tax_no = pmg_meals_pay
                 pmg_give_total = pmg_tax_yes + pmg_tax_no
		 
		         'if incom_family_cnt = 0 then
                       incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + incom_old + 1 '부양가족은 본인포함으로
                 'end if
			
			     '근로소득 간이세액 산출
                 inc_st_amt = 0  
                 inc_incom = 0
                 
				 Sql = "SELECT * FROM pay_income_amount where ('"&incom_month_amount&"' BETWEEN inc_from_amt and inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
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
		                epi_amt = pmg_give_tot * (epi_emp / 100)
                        epi_amt = int(epi_amt)
                        de_epi_amt = (int(epi_amt / 10)) * 10
                    else
		                de_epi_amt = 0
                 end if

                 '지방소득세
                 we_tax = inc_incom * (10 / 100)
                 we_tax = int(we_tax)
                 de_wetax = (int(we_tax / 10)) * 10 

                 de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
                 pmg_curr_pay = pmg_give_total - de_deduct_total
	      end if

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
   
		sql="insert into pay_month_give (pmg_yymm,pmg_id,pmg_emp_no,pmg_company,pmg_date,pmg_in_date,pmg_emp_name,pmg_emp_type,pmg_org_code,pmg_org_name,pmg_bonbu,pmg_saupbu,pmg_team,pmg_reside_place,pmg_reside_company,pmg_grade,pmg_position,pmg_base_pay,pmg_meals_pay,pmg_postage_pay,pmg_re_pay,pmg_overtime_pay,pmg_car_pay,pmg_position_pay,pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay,pmg_disabled_pay,pmg_family_pay,pmg_school_pay,pmg_qual_pay,pmg_other_pay1,pmg_other_pay2,pmg_other_pay3,pmg_tax_yes,pmg_tax_no,pmg_tax_reduced,pmg_give_total,pmg_bank_name,pmg_account_no,pmg_account_holder,cost_group,cost_center,pmg_reg_date,pmg_reg_user) values ('"&pmg_yymm_to&"','1','"&emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_in_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&pmg_base_pay&"','"&pmg_meals_pay&"','"&pmg_postage_pay&"','"&pmg_re_pay&"','"&pmg_overtime_pay&"','"&pmg_car_pay&"','"&pmg_position_pay&"','"&pmg_custom_pay&"','"&pmg_job_pay&"','"&pmg_job_support&"','"&pmg_jisa_pay&"','"&pmg_long_pay&"','"&pmg_disabled_pay&"','"&pmg_family_pay&"','"&pmg_school_pay&"','"&pmg_qual_pay&"','"&pmg_other_pay1&"','"&pmg_other_pay2&"','"&pmg_other_pay3&"','"&pmg_tax_yes&"','"&pmg_tax_no&"','"&pmg_tax_reduced&"','"&pmg_give_total&"','"&pmg_bank_name&"','"&pmg_account_no&"','"&pmg_account_holder&"','"&cost_group&"','"&cost_center&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
		
		sql="insert into pay_month_deduct (de_yymm,de_id,de_emp_no,de_company,de_date,de_emp_name,de_emp_type,de_org_code,de_org_name,de_bonbu,de_saupbu,de_team,de_reside_place,de_reside_company,de_grade,de_position,de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax,de_year_incom_tax2,de_year_wetax2,de_other_amt1,de_saving_amt,de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total,cost_group,cost_center,de_reg_date,de_reg_user) values ('"&pmg_yymm_to&"','1','"&emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&de_nps_amt&"','"&de_nhis_amt&"','"&de_epi_amt&"','"&de_longcare_amt&"','"&de_income_tax&"','"&de_wetax&"','"&de_year_incom_tax&"','"&de_year_wetax&"','"&de_year_incom_tax2&"','"&de_year_wetax2&"','"&de_other_amt1&"','"&de_saving_amt&"','"&de_sawo_amt&"','"&de_johab_amt&"','"&de_hyubjo_amt&"','"&de_school_amt&"','"&de_nhis_bla_amt&"','"&de_long_bla_amt&"','"&de_deduct_total&"','"&cost_group&"','"&cost_center&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
	   
		Rs.MoveNext()
    loop		
		response.write"<script language=javascript>"
		response.write"alert('전월 급여로 당월급여 기초 데이터가 만들어 졌습니다...');"		
		response.write"location.replace('insa_pay_month_batch.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('처리할 내역이 없습니다...');"		
		response.write"location.replace('insa_pay_month_batch.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
