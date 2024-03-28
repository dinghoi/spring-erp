<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
pmg_yymm_to=request("pmg_yymm_to")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

'if view_condi = "에스유에이치" then
'        v_company = "코리아디엔씨"
'   else
        v_company = view_condi
'end if

give_date = to_date '지급일

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) & "년 " & cstr(curr_mm) & "월 " & " 급여이월 내역서(개인별)-" & view_condi

savefilename = title_line &".xls"
'savefilename = "입사자 현황 -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"
'response.write(savefilename)

'당월 입사/퇴사일이 15일 이전이면 당월 급여대상임
st_es_date = mid(cstr(pmg_yymm_to),1,4) & "-" & mid(cstr(pmg_yymm_to),5,2) & "-" & "01"

st_in_date = mid(cstr(pmg_yymm_to),1,4) & "-" & mid(cstr(pmg_yymm_to),5,2) & "-" & "16"
rever_year = mid(cstr(pmg_yymm_to),1,4) '귀속년도

	sum_base_pay = 0
	sum_meals_pay = 0
	sum_postage_pay = 0
	sum_re_pay = 0
	sum_overtime_pay = 0
	sum_car_pay = 0
	sum_position_pay = 0
	sum_custom_pay = 0
	sum_job_pay = 0
	sum_job_support = 0
	sum_jisa_pay = 0
	sum_long_pay = 0
	sum_disabled_pay = 0
	sum_family_pay = 0
	sum_school_pay = 0
	sum_qual_pay = 0
	sum_other_pay1 = 0
	sum_other_pay2 = 0
	sum_other_pay3 = 0
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
    sum_nps_amt = 0
    sum_nhis_amt = 0
    sum_epi_amt = 0
    sum_longcare_amt = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
	sum_year_wetax = 0
	sum_year_incom_tax2 = 0
	sum_year_wetax2 = 0
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
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
           Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&v_company&"')  and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
end if

Rs.Open Sql, Dbconn, 1
'Response.write Sql
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">이월년월</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">귀속년월</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">지급일</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성  명</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">입사일</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">직급</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">본부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사업부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">팀</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">부서</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">기본급</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">식대</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">통신비</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">소급급여</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">연장근로수당</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">주차지원금</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">직책수당</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">고객관리수당</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">직무보조비</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">업무장려비</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">본지사근무비</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">근속수당</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">장애인수당</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">지급합계</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">국민연금</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">건강보험</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">고용보험</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">장기요양보험료</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">소득세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">지방소득세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">연말정산소득세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">연말정산지방세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">연말재정산소득세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">연말재정산지방세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">기타공제</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">사우회 회비</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">학자금상환</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">건강보험료정산</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">장기요양보험료정산</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">협조비</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">공제합계</div></td>  
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">차인지급액</div></td>  
  </tr>
    <%
		do until rs.eof 
		
		  emp_no = rs("emp_no")
		  emp_company = rs("emp_company")
		  emp_name = rs("emp_name")
		  emp_in_date = rs("emp_in_date")
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
		  
		  sql = "select * from pay_month_give where (pmg_yymm = '"&pmg_yymm&"' ) and (pmg_id = '1') and (pmg_emp_no = '"&emp_no&"') and (pmg_company = '"&emp_company&"')"
		  Set Rs_give = DbConn.Execute(SQL)
	      if not Rs_give.eof then	

		         pmg_company = Rs_give("pmg_company")
		         pmg_bonbu = Rs_give("pmg_bonbu")
		         pmg_saupbu = Rs_give("pmg_saupbu")
		         pmg_team = Rs_give("pmg_team")
		         pmg_org_name = Rs_give("pmg_org_name")
				 
		         pay_count = pay_count + 1
					  
		         pmg_base_pay = int(Rs_give("pmg_base_pay"))
	             pmg_meals_pay = int(Rs_give("pmg_meals_pay"))
	             pmg_postage_pay = int(Rs_give("pmg_postage_pay"))
	             pmg_re_pay = int(Rs_give("pmg_re_pay"))
	             pmg_overtime_pay = int(Rs_give("pmg_overtime_pay"))
	             pmg_car_pay = int(Rs_give("pmg_car_pay"))
                 pmg_position_pay = int(Rs_give("pmg_position_pay"))
	             pmg_custom_pay = int(Rs_give("pmg_custom_pay"))
	             pmg_job_pay = int(Rs_give("pmg_job_pay"))
	             pmg_job_support = int(Rs_give("pmg_job_support"))
	             pmg_jisa_pay = int(Rs_give("pmg_jisa_pay"))
	             pmg_long_pay = int(Rs_give("pmg_long_pay"))
	             pmg_disabled_pay = int(Rs_give("pmg_disabled_pay"))
	             pmg_give_total = int(Rs_give("pmg_give_total"))

	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_yymm_to%></div></td>
    <td width="110"><div align="center" class="style1"><%=give_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_no%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_grade%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_company%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_bonbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_saupbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_team%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_org_name%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_base_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_meals_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_postage_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_re_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_overtime_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_car_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_position_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_custom_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_job_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_job_support,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_jisa_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_long_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_disabled_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_give_total,0)%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>    
    <%
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
                        de_sawo_amt = int(Rs_dct("de_sawo_amt"))
                        de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                        de_school_amt = int(Rs_dct("de_school_amt"))
                        de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                        de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
                        de_deduct_tot = int(Rs_dct("de_deduct_total"))	
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
                        de_sawo_amt = 0
                        de_hyubjo_amt = 0
                        de_school_amt = 0
                        de_nhis_bla_amt = 0
                        de_long_bla_amt = 0
                        de_deduct_tot = 0
                 end if
                 Rs_dct.close()
				 
				 pmg_curr_pay = pmg_give_total - de_deduct_tot
							  
    %>    
    
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nps_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_longcare_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_other_amt1,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_sawo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_school_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_long_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_hyubjo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_deduct_tot,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_curr_pay,0)%></div></td>
  </tr>

	<%
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
                 de_sawo_amt = 0
                 de_hyubjo_amt = 0
                 de_school_amt = 0
                 de_nhis_bla_amt = 0
                 de_long_bla_amt = 0
                 de_deduct_tot = 0
				 
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

                 de_deduct_tot = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
                 pmg_curr_pay = pmg_give_total - de_deduct_tot
	%>		
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_yymm_to%></div></td>
    <td width="110"><div align="center" class="style1"><%=give_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_no%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_grade%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_company%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_bonbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_saupbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_team%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_org_name%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_base_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_meals_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_postage_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_re_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_overtime_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_car_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_position_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_custom_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_job_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_job_support,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_jisa_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_long_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_disabled_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_give_total,0)%></div></td>

    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nps_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_longcare_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_other_amt1,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_sawo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_school_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_long_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_hyubjo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_deduct_tot,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_curr_pay,0)%></div></td>
  </tr>
	<%		    				 
	      end if
		  Rs_give.close()

		  Rs.MoveNext()
	loop
			
	%>			

</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
