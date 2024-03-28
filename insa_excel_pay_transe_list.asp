<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")
in_empno = request("in_empno") 


curr_date = datevalue(mid(cstr(now()),1,10))

give_date = to_date '지급일

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여체크 내역서(개인별)"

savefilename = title_line +".xls"
'savefilename = "입사자 현황 -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"
'response.write(savefilename)

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
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'당월 입사일이 15일 이전이면 당월 급여대상임
st_es_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + "01"
st_in_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + "16"

if condi = "" then
      Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"')  and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
   else  
      if owner_view = "C" then 
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_name like '%"&condi&"%') ORDER BY emp_in_date,emp_no ASC"
         else
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_no = '"&condi&"') ORDER BY emp_in_date,emp_no ASC"
	  end if
end if
Rs_emp.Open Sql, Dbconn, 1

'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
'Rs.Open Sql, Dbconn, 1

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
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="16" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="9" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">인적사항</div></td>
    <td colspan="14" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">기본급여 및 제수당</div></td>
    <td colspan="14" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">공제 및 차인지급액</div></td>
  </tr>
  <tr>
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
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">연말정산지방소득세</div></td>
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
		do until Rs_emp.eof 
		
		  emp_no = Rs_emp("emp_no")
		  emp_name = Rs_emp("emp_name")
		  emp_grade = Rs_emp("emp_grade")
		  emp_position = Rs_emp("emp_position")
		  emp_in_date = rs_emp("emp_in_date")
		  
		  Sql = "SELECT * FROM pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_emp_no = '"&emp_no&"' and pmg_id = '1' and (pmg_company = '"+view_condi+"')"
          Set rs_give = DbConn.Execute(SQL)
		  if not rs_give.eof then
				pmg_org_code = rs_give("pmg_org_code")
				pmg_org_name = rs_give("pmg_org_name")
	            pmg_company = rs_give("pmg_company")
				pmg_bonbu = rs_give("pmg_bonbu")
				pmg_saupbu = rs_give("pmg_saupbu")
				pmg_team = rs_give("pmg_team")
				pmg_org_name = rs_give("pmg_org_name")
				pmg_reside_place = rs_give("pmg_reside_place")
				pmg_reside_company = rs_give("pmg_reside_company")
	            pmg_emp_type = rs_give("pmg_emp_type")
	            pmg_grade = rs_give("pmg_grade")
            	pmg_position = rs_give("pmg_position")
				
	            pmg_base_pay = int(rs_give("pmg_base_pay"))
	            pmg_meals_pay = int(rs_give("pmg_meals_pay"))
             	pmg_postage_pay = int(rs_give("pmg_postage_pay"))
	            pmg_re_pay = int(rs_give("pmg_re_pay"))
	            pmg_overtime_pay = int(rs_give("pmg_overtime_pay"))
	            pmg_car_pay = int(rs_give("pmg_car_pay"))
	            pmg_position_pay = int(rs_give("pmg_position_pay"))
	            pmg_custom_pay = int(rs_give("pmg_custom_pay"))
	            pmg_job_pay = int(rs_give("pmg_job_pay"))
	            pmg_job_support = int(rs_give("pmg_job_support"))
	            pmg_jisa_pay = int(rs_give("pmg_jisa_pay"))
	            pmg_long_pay = int(rs_give("pmg_long_pay"))
	            pmg_disabled_pay = int(rs_give("pmg_disabled_pay"))
	            pmg_family_pay = int(rs_give("pmg_family_pay"))
	            pmg_school_pay = int(rs_give("pmg_school_pay"))
	            pmg_qual_pay = int(rs_give("pmg_qual_pay"))
	            pmg_other_pay1 = int(rs_give("pmg_other_pay1"))
	            pmg_other_pay2 = int(rs_give("pmg_other_pay2"))
	            pmg_other_pay3 = int(rs_give("pmg_other_pay3"))
	            pmg_tax_yes = int(rs_give("pmg_tax_yes"))
	            pmg_tax_no = int(rs_give("pmg_tax_no"))
	            pmg_tax_reduced = int(rs_give("pmg_tax_reduced"))
			    pmg_give_total = int(rs_give("pmg_give_total"))
	         else
                pmg_org_code = ""
				pmg_org_name = ""
	            pmg_company = ""
				pmg_bonbu = ""
				pmg_saupbu = ""
				pmg_team = ""
				pmg_org_name = ""
				pmg_reside_place = ""
				pmg_reside_company = ""
	            pmg_emp_type = ""
	            pmg_grade = ""
            	pmg_position = ""
	
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
          end if
          rs_give.close()
		  
		  pay_count = pay_count + 1
					  
		  sum_base_pay = sum_base_pay + pmg_base_pay
	      sum_meals_pay = sum_meals_pay + pmg_meals_pay
	      sum_postage_pay = sum_postage_pay + pmg_postage_pay
	      sum_re_pay = sum_re_pay + pmg_re_pay
	      sum_overtime_pay = sum_overtime_pay + pmg_overtime_pay
	      sum_car_pay = sum_car_pay + pmg_car_pay
          sum_position_pay = sum_position_pay + pmg_position_pay
	      sum_custom_pay = sum_custom_pay + pmg_custom_pay
	      sum_job_pay = sum_job_pay + pmg_job_pay
	      sum_job_support = sum_job_support + pmg_job_support
	      sum_jisa_pay = sum_jisa_pay + pmg_jisa_pay
	      sum_long_pay = sum_long_pay + pmg_long_pay
	      sum_disabled_pay = sum_disabled_pay + pmg_disabled_pay
	      sum_give_tot = sum_give_tot + pmg_give_total
		  
	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=to_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_no%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_grade%></div></td>
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
    <%
	      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
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
							  
	 	   sum_nps_amt = sum_nps_amt + de_nps_amt
           sum_nhis_amt = sum_nhis_amt + de_nhis_amt
           sum_epi_amt = sum_epi_amt + de_epi_amt
		   sum_longcare_amt = sum_longcare_amt + de_longcare_amt
           sum_income_tax = sum_income_tax + de_income_tax
           sum_wetax = sum_wetax + de_wetax
		   sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
           sum_year_wetax = sum_year_wetax + de_year_wetax
           sum_other_amt1 = sum_other_amt1 + de_other_amt1
           sum_sawo_amt = sum_sawo_amt + de_sawo_amt
           sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
           sum_school_amt = sum_school_amt + de_school_amt
           sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
           sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
		   sum_deduct_tot = sum_deduct_tot + de_deduct_tot
							  
    %>    
    
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nps_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_longcare_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax,0)%></div></td>
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
	    Rs_emp.MoveNext()
	loop
	
	sum_curr_pay = sum_give_tot - sum_deduct_tot
	
	%>
    
  <tr>    
    <th colspan="11" style=" border-top:1px solid #e3e3e3;"><div align="center" class="style1">총계</div></th>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_base_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_meals_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_postage_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_re_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sumpmg_car_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_position_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_custom_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_support,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_long_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_give_tot,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_nps_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_nhis_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_epi_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_longcare_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_income_tax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_wetax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_other_amt1,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_sawo_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_school_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_nhis_bla_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_long_bla_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_hyubjo_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_deduct_tot,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
  </tr>
</table>
</body>
</html>
<%
Rs_emp.Close()
Set Rs_emp = Nothing
%>
