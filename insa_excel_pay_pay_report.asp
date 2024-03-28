<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
pmg_emp_name=request("pmg_emp_name")

curr_date = datevalue(mid(cstr(now()),1,10))

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여이월 내역서(개인별)"

savefilename = title_line +".xls"

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
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') "
if view_condi <> "전체" then
	Sql = Sql & " and (pmg_company = '"+view_condi+"') "
end if
If Trim(pmg_emp_name&"")<>"" Then
	Sql = Sql & " and pmg_emp_name like '%" & pmg_emp_name & "%' "
End If
Sql = Sql & " ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"

Rs.Open Sql, Dbconn, 1

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
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">귀속년월</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성  명</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">입사일</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">직급</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">본부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사업부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">팀</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">부서</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">상주처</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">상주처회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">비용센타그룹</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">비용구분</div></td>
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
		
		  emp_no = rs("pmg_emp_no")
		  pmg_company = rs("pmg_company")
		  pmg_give_tot = rs("pmg_give_total")
		  pay_count = pay_count + 1
					  
		  sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
	      sum_meals_pay = sum_meals_pay + int(rs("pmg_meals_pay"))
	      sum_postage_pay = sum_postage_pay + int(rs("pmg_postage_pay"))
	      sum_re_pay = sum_re_pay + int(rs("pmg_re_pay"))
	      sum_overtime_pay = sum_overtime_pay + int(rs("pmg_overtime_pay"))
	      sum_car_pay = sum_car_pay + int(rs("pmg_car_pay"))
          sum_position_pay = sum_position_pay + int(rs("pmg_position_pay"))
	      sum_custom_pay = sum_custom_pay + int(rs("pmg_custom_pay"))
	      sum_job_pay = sum_job_pay + int(rs("pmg_job_pay"))
	      sum_job_support = sum_job_support + int(rs("pmg_job_support"))
	      sum_jisa_pay = sum_jisa_pay + int(rs("pmg_jisa_pay"))
	      sum_long_pay = sum_long_pay + int(rs("pmg_long_pay"))
	      sum_disabled_pay = sum_disabled_pay + int(rs("pmg_disabled_pay"))
	      sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))
		  
	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_emp_no")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_emp_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_in_date")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_grade")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_company")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_bonbu")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_saupbu")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_team")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_org_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_reside_place")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_reside_company")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("cost_group")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("cost_center")%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_base_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_meals_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_postage_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_re_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_overtime_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_car_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_position_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_custom_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_job_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_job_support"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_jisa_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_long_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_disabled_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_give_total"),0)%></div></td>
    <%
	      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+pmg_company+"')"
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
		   
		   pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  
	 	   sum_nps_amt = sum_nps_amt + de_nps_amt
           sum_nhis_amt = sum_nhis_amt + de_nhis_amt
           sum_epi_amt = sum_epi_amt + de_epi_amt
		   sum_longcare_amt = sum_longcare_amt + de_longcare_amt
           sum_income_tax = sum_income_tax + de_income_tax
           sum_wetax = sum_wetax + de_wetax
		   sum_income_tax = sum_income_tax + de_income_tax
           sum_wetax = sum_wetax + de_wetax
		   sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
           sum_year_wetax = sum_year_wetax + de_year_wetax
		   sum_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
           sum_year_wetax2 = sum_year_wetax2 + de_year_wetax2
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
		Rs.MoveNext()
	loop
	
	sum_curr_pay = sum_give_tot - sum_deduct_tot
	
	%>
  <tr valign="middle" class="style11">
    <td colspan="13" width="110"><div align="center" class="style1">총계</div></td>
    <td width="110"><div align="center" class="style1"><%=formatnumber(pay_count,0)%>&nbsp;명</div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_base_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_meals_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_postage_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_re_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_car_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_position_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_custom_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_job_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_job_support,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_long_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_give_total,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_nps_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_nhis_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_longcare_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_year_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_year_wetax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_other_amt1,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_sawo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_school_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_nhis_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_long_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_hyubjo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_deduct_tot,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
  </tr>    
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
