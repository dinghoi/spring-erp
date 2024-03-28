<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'상주회사 기준 엑셀

Dim Rs
Dim stay_name

dim com_tab(6)
dim pay_count(6)
dim sum_base_pay(6)
dim sum_meals_pay(6)
dim sum_postage_pay(6)
dim sum_re_pay(6)
dim sum_overtime_pay(6)
dim sum_car_pay(6)
dim sum_position_pay(6)
dim sum_custom_pay(6)
dim sum_job_pay(6)
dim sum_job_support(6)
dim sum_jisa_pay(6)
dim sum_long_pay(6)
dim sum_disabled_pay(6)
dim sum_give_tot(6)

dim sum_nps_amt(6)
dim sum_nhis_amt(6)
dim sum_epi_amt(6)
dim sum_longcare_amt(6)
dim sum_income_tax(6)
dim sum_wetax(6)
dim sum_year_incom_tax(6)
dim sum_year_wetax(6)
dim sum_other_amt1(6)
dim sum_sawo_amt(6)
dim sum_hyubjo_amt(6)
dim sum_school_amt(6)
dim sum_nhis_bla_amt(6)
dim sum_long_bla_amt(6)
dim sum_deduct_tot(6)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")

curr_date = datevalue(mid(cstr(now()),1,10))

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여대장(Cost Center)"

savefilename = title_line +".xls"
'savefilename = "입사자 현황 -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"
'response.write(savefilename)

	for i = 1 to 6
        com_tab(i) = ""
        pay_count(i) = 0
        sum_base_pay(i) = 0
        sum_meals_pay(i) = 0
        sum_postage_pay(i) = 0
        sum_re_pay(i) = 0
        sum_overtime_pay(i) = 0
        sum_car_pay(i) = 0
        sum_position_pay(i) = 0
        sum_custom_pay(i) = 0
        sum_job_pay(i) = 0
        sum_job_support(i) = 0
        sum_jisa_pay(i) = 0
        sum_long_pay(i) = 0
        sum_disabled_pay(i) = 0
        sum_give_tot(i) = 0
        sum_nps_amt(i) = 0
        sum_nhis_amt(i) = 0
        sum_epi_amt(i) = 0
        sum_longcare_amt(i) = 0
        sum_income_tax(i) = 0
        sum_wetax(i) = 0
        sum_year_incom_tax(i) = 0
        sum_year_wetax(i) = 0
        sum_other_amt1(i) = 0
        sum_sawo_amt(i) = 0
        sum_hyubjo_amt(i) = 0
        sum_school_amt(i) = 0
        sum_nhis_bla_amt(i) = 0
        sum_long_bla_amt(i) = 0
        sum_deduct_tot(i) = 0
    next
	
	sum_curr_pay = 0	
	
	org_base_pay = 0
	org_meals_pay = 0
	org_postage_pay = 0
	org_re_pay = 0
	org_overtime_pay = 0
	org_car_pay = 0
	org_position_pay = 0
	org_custom_pay = 0
	org_job_pay = 0
	org_job_support = 0
	org_jisa_pay = 0
	org_long_pay = 0
	org_disabled_pay = 0
	org_give_tot = 0	
	org_count = 0
	
	sap_base_pay = 0
	sap_meals_pay = 0
	sap_postage_pay = 0
	sap_re_pay = 0
	sap_overtime_pay = 0
	sap_car_pay = 0
	sap_position_pay = 0
	sap_custom_pay = 0
	sap_job_pay = 0
	sap_job_support = 0
	sap_jisa_pay = 0
	sap_long_pay = 0
	sap_disabled_pay = 0
	sap_give_tot = 0	
	sap_count = 0

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

order_Sql = " ORDER BY pmg_reside_company,pmg_saupbu,pmg_org_name,pmg_emp_no ASC"
'order_Sql = " ORDER BY pmg_org_name,pmg_emp_no ASC"
if view_condi = "전체" then
      com_tab(1) = "케이원정보통신"
	  com_tab(2) = "휴디스"
	  com_tab(3) = "케이네트웍스"
	  com_tab(4) = "에스유에이치"
	  com_tab(5) = "코리아디엔씨"
	  com_tab(6) = "합계"
	  where_sql = " WHERE (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1')" 
   else  
      com_tab(1) = view_condi
	  com_tab(6) = "합계"
	  where_sql = " WHERE (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if   

sql = "select * from pay_month_give " + where_sql + order_sql
'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
	pmg_company = rs("pmg_company")
    				  
    for i = 1 to 6
        if com_tab(i) = rs("pmg_company") then
	             pay_count(i) = pay_count(i) + 1
				 pay_count(6) = pay_count(6) + 1
		         sum_base_pay(i) = sum_base_pay(i) + int(rs("pmg_base_pay"))
                 sum_meals_pay(i) = sum_meals_pay(i) + int(rs("pmg_meals_pay"))
                 sum_postage_pay(i) = sum_postage_pay(i) + int(rs("pmg_postage_pay"))
                 sum_re_pay(i) = sum_re_pay(i) + int(rs("pmg_re_pay"))
                 sum_overtime_pay(i) = sum_overtime_pay(i) + int(rs("pmg_overtime_pay"))
                 sum_car_pay(i) = sum_car_pay(i) + int(rs("pmg_car_pay"))
                 sum_position_pay(i) = sum_position_pay(i) + int(rs("pmg_position_pay"))
                 sum_custom_pay(i) = sum_custom_pay(i) + int(rs("pmg_custom_pay"))
                 sum_job_pay(i) = sum_job_pay(i) + int(rs("pmg_job_pay"))
                 sum_job_support(i) = sum_job_support(i) + int(rs("pmg_job_support"))
                 sum_jisa_pay(i) = sum_jisa_pay(i) + int(rs("pmg_jisa_pay"))
                 sum_long_pay(i) = sum_long_pay(i) + int(rs("pmg_long_pay"))
                 sum_disabled_pay(i) = sum_disabled_pay(i) + int(rs("pmg_disabled_pay"))
                 sum_give_tot(i) = sum_give_tot(i) + int(rs("pmg_give_total"))
				 
				 sum_base_pay(6) = sum_base_pay(6) + int(rs("pmg_base_pay"))
                 sum_meals_pay(6) = sum_meals_pay(6) + int(rs("pmg_meals_pay"))
                 sum_postage_pay(6) = sum_postage_pay(6) + int(rs("pmg_postage_pay"))
                 sum_re_pay(6) = sum_re_pay(6) + int(rs("pmg_re_pay"))
                 sum_overtime_pay(6) = sum_overtime_pay(6) + int(rs("pmg_overtime_pay"))
                 sum_car_pay(6) = sum_car_pay(6) + int(rs("pmg_car_pay"))
                 sum_position_pay(6) = sum_position_pay(6) + int(rs("pmg_position_pay"))
                 sum_custom_pay(6) = sum_custom_pay(6) + int(rs("pmg_custom_pay"))
                 sum_job_pay(6) = sum_job_pay(6) + int(rs("pmg_job_pay"))
                 sum_job_support(6) = sum_job_support(6) + int(rs("pmg_job_support"))
                 sum_jisa_pay(6) = sum_jisa_pay(6) + int(rs("pmg_jisa_pay"))
                 sum_long_pay(6) = sum_long_pay(6) + int(rs("pmg_long_pay"))
                 sum_disabled_pay(6) = sum_disabled_pay(6) + int(rs("pmg_disabled_pay"))
                 sum_give_tot(6) = sum_give_tot(6) + int(rs("pmg_give_total"))
	    end if		 	 
	next		

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
	 
     for i = 1 to 6
        if com_tab(i) = rs("pmg_company") then
		         sum_nps_amt(i) = sum_nps_amt(i) + de_nps_amt
                 sum_nhis_amt(i) = sum_nhis_amt(i) + de_nhis_amt
                 sum_epi_amt(i) = sum_epi_amt(i) + de_epi_amt
	             sum_longcare_amt(i) = sum_longcare_amt(i) + de_longcare_amt
                 sum_income_tax(i) = sum_income_tax(i) + de_income_tax
                 sum_wetax(i) = sum_wetax(i) + de_wetax
	             sum_year_incom_tax(i) = sum_year_incom_tax(i) + de_year_incom_tax
                 sum_year_wetax(i) = sum_year_wetax(i) + de_year_wetax
                 sum_other_amt1(i) = sum_other_amt1(i) + de_other_amt1
                 sum_sawo_amt(i) = sum_sawo_amt(i) + de_sawo_amt
                 sum_hyubjo_amt(i) = sum_hyubjo_amt(i) + de_hyubjo_amt
                 sum_school_amt(i) = sum_school_amt(i) + de_school_amt
                 sum_nhis_bla_amt(i) = sum_nhis_bla_amt(i) + de_nhis_bla_amt
                 sum_long_bla_amt(i) = sum_long_bla_amt(i) + de_long_bla_amt
	             sum_deduct_tot(i) = sum_deduct_tot(i) + de_deduct_tot
				 
				 sum_nps_amt(6) = sum_nps_amt(6) + de_nps_amt
                 sum_nhis_amt(6) = sum_nhis_amt(6) + de_nhis_amt
                 sum_epi_amt(6) = sum_epi_amt(6) + de_epi_amt
	             sum_longcare_amt(6) = sum_longcare_amt(6) + de_longcare_amt
                 sum_income_tax(6) = sum_income_tax(6) + de_income_tax
                 sum_wetax(6) = sum_wetax(6) + de_wetax
	             sum_year_incom_tax(6) = sum_year_incom_tax(6) + de_year_incom_tax
                 sum_year_wetax(6) = sum_year_wetax(6) + de_year_wetax
                 sum_other_amt1(6) = sum_other_amt1(6) + de_other_amt1
                 sum_sawo_amt(6) = sum_sawo_amt(6) + de_sawo_amt
                 sum_hyubjo_amt(6) = sum_hyubjo_amt(6) + de_hyubjo_amt
                 sum_school_amt(6) = sum_school_amt(6) + de_school_amt
                 sum_nhis_bla_amt(6) = sum_nhis_bla_amt(6) + de_nhis_bla_amt
                 sum_long_bla_amt(6) = sum_long_bla_amt(6) + de_long_bla_amt
	             sum_deduct_tot(6) = sum_deduct_tot(6) + de_deduct_tot
	    end if		 	 
	 next		

	rs.movenext()
loop
rs.close()

f_sw = "1"
reside_sw = "1"
bigo_org = ""
bigo_reside_company = ""

sql = "select * from pay_month_give " + where_sql + order_sql 
'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"

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
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="16" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="8" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">인적사항</div></td>
    <td colspan="13" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">기본급여 및 제수당</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">지급액</div></td>
  </tr>
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">귀속년월</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성명</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사업부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">부서</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">직급</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">급여성격</div></td>
    
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
  </tr>
    <%
       do until rs.eof

          if rs("pmg_saupbu") = "" or isnull(rs("pmg_saupbu")) then
		          pmg_saupbu = ""
			 else
			      pmg_saupbu = rs("pmg_saupbu")
		  end if
		  if rs("pmg_reside_company") = "" or isnull(rs("pmg_reside_company")) then
		          pmg_reside_company = ""
			 else
			      pmg_reside_company = rs("pmg_reside_company")
		  end if
		  
		  if f_sw = "1" then
		        bigo_org = pmg_saupbu
				bigo_reside_company = pmg_reside_company
				f_sw = "2"
		  end if
          
          pmg_base_pay = int(rs("pmg_base_pay"))
          pmg_meals_pay = int(rs("pmg_meals_pay"))
       	  pmg_postage_pay = int(rs("pmg_postage_pay"))
	      pmg_re_pay = int(rs("pmg_re_pay"))
	      pmg_overtime_pay = int(rs("pmg_overtime_pay"))
	      pmg_car_pay = int(rs("pmg_car_pay"))
	      pmg_position_pay = int(rs("pmg_position_pay"))
	      pmg_custom_pay = int(rs("pmg_custom_pay"))
	      pmg_job_pay = int(rs("pmg_job_pay"))
	      pmg_job_support = int(rs("pmg_job_support"))
	      pmg_jisa_pay = int(rs("pmg_jisa_pay"))
	      pmg_long_pay = int(rs("pmg_long_pay"))
	      pmg_disabled_pay = int(rs("pmg_disabled_pay"))
		  pmg_give_total = int(rs("pmg_give_total"))
		
		  if reside_sw = "1" then
		    if pmg_reside_company = "" then
		      if bigo_org <> pmg_saupbu then
	 %>
  <tr valign="middle" class="style11">
    <td colspan="4" style="background:#E0FFFF;"><div align="center" class="style1"><%=bigo_org%></div></td>
    <th colspan="3" style="background:#E0FFFF;"><div align="center" class="style1">사업부계</div></th>
    <td width="120" style="background:#E0FFFF;"><div align="center" class="style1"><%=sap_count%>&nbsp;명</div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_base_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_meals_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_postage_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_re_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_overtime_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_car_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_position_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_custom_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_job_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_job_support,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_jisa_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_long_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_disabled_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_give_tot,0)%></div></td>
  </tr>    
     <%
	                bigo_org = pmg_saupbu
		   
				   	sap_base_pay = 0
	                sap_meals_pay = 0
	                sap_postage_pay = 0
	                sap_re_pay = 0
	                sap_overtime_pay = 0
	                sap_car_pay = 0
	                sap_position_pay = 0
	                sap_custom_pay = 0
	                sap_job_pay = 0
	                sap_job_support = 0
	                sap_jisa_pay = 0
	                sap_long_pay = 0
	                sap_disabled_pay = 0
	                sap_give_tot = 0	
					sap_count = 0	
			   end if
			 else
			        reside_sw = "2"
	 %>
  <tr valign="middle" class="style11">
    <td colspan="4" style="background:#E0FFFF;"><div align="center" class="style1"><%=bigo_org%></div></td>
    <th colspan="3" style="background:#E0FFFF;"><div align="center" class="style1">사업부계</div></th>
    <td width="120" style="background:#E0FFFF;"><div align="center" class="style1"><%=sap_count%>&nbsp;명</div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_base_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_meals_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_postage_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_re_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_overtime_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_car_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_position_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_custom_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_job_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_job_support,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_jisa_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_long_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_disabled_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(sap_give_tot,0)%></div></td>
  </tr>                 
     <%		   
				   	sap_base_pay = 0
	                sap_meals_pay = 0
	                sap_postage_pay = 0
	                sap_re_pay = 0
	                sap_overtime_pay = 0
	                sap_car_pay = 0
	                sap_position_pay = 0
	                sap_custom_pay = 0
	                sap_job_pay = 0
	                sap_job_support = 0
	                sap_jisa_pay = 0
	                sap_long_pay = 0
	                sap_disabled_pay = 0
	                sap_give_tot = 0	
					sap_count = 0	
		    end if	
		  end if 
	 %>			
     <%	     
		  'if bigo_org <> pmg_saupbu then
		  if bigo_reside_company <> pmg_reside_company then
	 %>
  <tr valign="middle" class="style11">
    <td colspan="4" style="background:#E0FFFF;"><div align="center" class="style1"><%=bigo_reside_company%></div></td>
    <th colspan="3" style="background:#E0FFFF;"><div align="center" class="style1">상주회사계</div></th>
    <td width="120" style="background:#E0FFFF;"><div align="center" class="style1"><%=org_count%>&nbsp;명</div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_base_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_meals_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_postage_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_re_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_overtime_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_car_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_position_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_custom_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_job_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_job_support,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_jisa_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_long_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_disabled_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_give_tot,0)%></div></td>
  </tr>    		  
	<%             
	                bigo_org = pmg_saupbu
				    bigo_reside_company = pmg_reside_company
				   
				   	org_base_pay = 0
	                org_meals_pay = 0
	                org_postage_pay = 0
	                org_re_pay = 0
	                org_overtime_pay = 0
	                org_car_pay = 0
	                org_position_pay = 0
	                org_custom_pay = 0
	                org_job_pay = 0
	                org_job_support = 0
	                org_jisa_pay = 0
	                org_long_pay = 0
	                org_disabled_pay = 0
	                org_give_tot = 0	
					org_count = 0	
			end if
	        
			org_count = org_count + 1
		    org_base_pay = org_base_pay + pmg_base_pay
	        org_meals_pay = org_meals_pay + pmg_meals_pay
	        org_postage_pay = org_postage_pay + pmg_postage_pay
	        org_re_pay = org_re_pay + pmg_re_pay
	        org_overtime_pay = org_overtime_pay + pmg_overtime_pay
	        org_car_pay = org_car_pay + pmg_car_pay
            org_position_pay = org_position_pay + pmg_position_pay
	        org_custom_pay = org_custom_pay + pmg_custom_pay
	        org_job_pay = org_job_pay + pmg_job_pay
	        org_job_support = org_job_support + pmg_job_support
	        org_jisa_pay = org_jisa_pay + pmg_jisa_pay
	        org_long_pay = org_long_pay + pmg_long_pay
	        org_disabled_pay = org_disabled_pay + pmg_disabled_pay
	        org_give_tot = org_give_tot + pmg_give_total
			
			sap_count = sap_count + 1
		    sap_base_pay = sap_base_pay + pmg_base_pay
	        sap_meals_pay = sap_meals_pay + pmg_meals_pay
	        sap_postage_pay = sap_postage_pay + pmg_postage_pay
	        sap_re_pay = sap_re_pay + pmg_re_pay
	        sap_overtime_pay = sap_overtime_pay + pmg_overtime_pay
	        sap_car_pay = sap_car_pay + pmg_car_pay
            sap_position_pay = sap_position_pay + pmg_position_pay
	        sap_custom_pay = sap_custom_pay + pmg_custom_pay
	        sap_job_pay = sap_job_pay + pmg_job_pay
	        sap_job_support = sap_job_support + pmg_job_support
	        sap_jisa_pay = sap_jisa_pay + pmg_jisa_pay
	        sap_long_pay = sap_long_pay + pmg_long_pay
	        sap_disabled_pay = sap_disabled_pay + pmg_disabled_pay
	        sap_give_tot = sap_give_tot + pmg_give_total
		  
	%>
  <tr valign="middle" class="style11">
    <td width="120"><div align="center" class="style1"><%=rs("pmg_yymm")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("pmg_emp_name")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("pmg_emp_no")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("pmg_company")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("pmg_saupbu")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("pmg_org_name")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("pmg_grade")%></div></td>
    <td width="120"><div align="center" class="style1"><%=rs("cost_center")%></div></td>

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
  </tr>
	<%  
	    rs.MoveNext()
	loop

	%>
  <tr valign="middle" class="style11">
    <td colspan="4" style="background:#E0FFFF;"><div align="center" class="style1"><%=bigo_org%></div></td>
    <th colspan="3" style="background:#E0FFFF;"><div align="center" class="style1">상주회사계</div></th>
    <td width="120" style="background:#E0FFFF;"><div align="center" class="style1"><%=org_count%>&nbsp;명</div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_base_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_meals_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_postage_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_re_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_overtime_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_car_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_position_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_custom_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_job_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_job_support,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_jisa_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_long_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_disabled_pay,0)%></div></td>
    <td width="100" style="background:#E0FFFF;"><div align="right" class="style1"><%=formatnumber(org_give_tot,0)%></div></td>
  </tr>    		 
      
  <tr>    
    <th colspan="8" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">총계</div></th>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_base_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_meals_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_postage_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_re_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_car_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_position_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_custom_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_support(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_long_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay(6),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_give_tot(6),0)%></div></td>
  </tr>
<%
    for i = 1 to 6 
        if	com_tab(i) <> "" then
%>	  
  <tr>    
    <th colspan="7" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1"><%=com_tab(i)%></div></th>
    <th width="120" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1"><%=pay_count(i)%>&nbsp;명</div></th>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_base_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_meals_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_postage_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_re_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_car_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_position_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_custom_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_support(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_long_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay(i),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_give_tot(i),0)%></div></td>
  </tr>  
<%
	     end if
    next
%>      
</table>
</body>
</html>
<%
rs.Close()
Set rs = Nothing
%>
