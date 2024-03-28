<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))
to_yyyy = mid(cstr(to_date),1,4)
to_mm = mid(cstr(to_date),6,2)
to_dd = mid(cstr(to_date),9,2)

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 사업부별 급여내역(" + view_condi + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

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
' 사업부	
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
	org_family_pay = 0
	org_school_pay = 0
	org_qual_pay = 0
	org_other_pay1 = 0
	org_other_pay2 = 0
	org_other_pay3 = 0
	org_tax_yes = 0
	org_tax_no = 0
	org_tax_reduced = 0
	org_give_tot = 0
    org_nps_amt = 0
    org_nhis_amt = 0
    org_epi_amt = 0
    org_longcare_amt = 0
    org_income_tax = 0
    org_wetax = 0
	org_year_incom_tax = 0
    org_year_wetax = 0
	org_year_incom_tax2 = 0
    org_year_wetax2 = 0
    org_other_amt1 = 0
    org_sawo_amt = 0
    org_hyubjo_amt = 0
    org_school_amt = 0
    org_nhis_bla_amt = 0
    org_long_bla_amt = 0
	org_deduct_tot = 0
	
	org_pay_count = 0	
	org_curr_pay = 0
	
' 팀	
	team_base_pay = 0 
	team_meals_pay = 0
	team_postage_pay = 0
	team_re_pay = 0
	team_overtime_pay = 0
	team_car_pay = 0
	team_position_pay = 0
	team_custom_pay = 0
	team_job_pay = 0
	team_job_support = 0
	team_jisa_pay = 0
	team_long_pay = 0
	team_disabled_pay = 0
	team_family_pay = 0
	team_school_pay = 0
	team_qual_pay = 0
	team_other_pay1 = 0
	team_other_pay2 = 0
	team_other_pay3 = 0
	team_tax_yes = 0
	team_tax_no = 0
	team_tax_reduced = 0
	team_give_tot = 0
    team_nps_amt = 0
    team_nhis_amt = 0
    team_epi_amt = 0
    team_longcare_amt = 0
    team_income_tax = 0
    team_wetax = 0
	team_year_incom_tax = 0
    team_year_wetax = 0
	team_year_incom_tax2 = 0
    team_year_wetax2 = 0
    team_other_amt1 = 0
    team_sawo_amt = 0
    team_hyubjo_amt = 0
    team_school_amt = 0
    team_nhis_bla_amt = 0
    team_long_bla_amt = 0
	team_deduct_tot = 0
	
	team_pay_count = 0	
	team_curr_pay = 0	
	
give_date = to_date '지급일

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

if view_condi = "전체" then
      Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') ORDER BY pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_code,pmg_emp_no ASC"
   else	  
	  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_code,pmg_emp_no ASC"
end if
Rs.Open Sql, Dbconn, 1


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
	</head>
	<body>
    	<div id="wrap">			 
			<div id="container">
                <h3 class="insa"><%=title_line%></h3>
				<div class="gView">
                <table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
				               <th colspan="2" height="30" scope="col" style=" border-bottom:1px solid #e3e3e3;">사업부&nbsp;명</th>
				               <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;">공제 및 차인지급액</th>
			                </tr>
                            <tr>
								<td colspan="2" height="30" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">&nbsp;</td> 
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">기본급</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">식대</td>  
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">차량유지비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">통신비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">소급급여</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">연장근로<br>수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">주차지원금</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">국민연금</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">건강보험</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">고용보험</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">장기요양<br>보험료</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">소득세</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">지방소득세</td>
							</tr>
                            <tr>
								<td colspan="2" height="30" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">&nbsp;</td> 
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">직책수당</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">고객관리<br>수당</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">직무보조비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">업무장려비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">본지사<br>근무비</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">근속수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">장애인수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">기타공제</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">사우회<br>회비</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">학자금상환</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">건강보험료<br>정산</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">장기요양<br>보험료정산</td>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">공제합계</th>
							</tr>
                            <tr>
								<td colspan="2" height="30" scope="col" style=" border-bottom:2px solid #515254; background:#f8f8f8;">&nbsp;</td> 
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <th scope="col" style=" border-bottom:2px solid #515254;">지급합계</th>
                                <td scope="col" style=" border-bottom:2px solid #515254;">협조비</td>
                                
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말정산<br>소득세</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말정산<br>지방세</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말재정산<br>소득세</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말재정산<br>지방세</td>
                                <th scope="col" style=" border-bottom:2px solid #515254; font-size:12px">차인지급액</th>
							</tr>
						</thead>
						<tbody>
			<%
                        if rs.eof or rs.bof then
		                         bi_org = ""
								 bi_company = ""
								 bi_bonbu = ""
			                     bi_team = ""
		                   else						  
			                     if isnull(rs("pmg_saupbu")) or rs("pmg_saupbu") = "" then	
				 	                      bi_org = ""
										  bi_company = rs("pmg_company")
										  bi_bonbu = rs("pmg_bonbu")
				                    else
					                      bi_org = rs("pmg_saupbu")
										  bi_company = rs("pmg_company")
										  bi_bonbu = rs("pmg_bonbu")
			                     end if
			                     if isnull(rs("pmg_team")) or rs("pmg_team") = "" then	
				 	                      bi_team = ""
				                    else
					                      bi_team = rs("pmg_team")
			                     end if
		                end if		

						do until rs.eof
						
						   if isnull(rs("pmg_saupbu")) or rs("pmg_saupbu") = "" then
				                   pmg_saupbu = ""
								   pmg_company = rs("pmg_company")
								   pmg_bonbu = rs("pmg_bonbu")
			                  else
			                       pmg_saupbu = rs("pmg_saupbu")
								   pmg_company = rs("pmg_company")
								   pmg_bonbu = rs("pmg_bonbu")
		                   end if
		                   if isnull(rs("pmg_team")) or rs("pmg_team") = "" then
		  	                       pmg_team = ""
	 	                      else
			                       pmg_team = rs("pmg_team")
		                   end if		

                      if bi_org <> pmg_saupbu then
		                    org_curr_pay = org_give_tot - org_deduct_tot
							
							if bi_org = "" and bi_bonbu = "" then
							        bi_org = bi_company 
							   else
							       if bi_org = "" then 
								       bi_org = bi_bonbu + " 직할"
								   end if
						    end if
	       %>
                              <tr>
                                <td rowspan="3" class="first"><%=bi_org%>&nbsp;&nbsp;계</td>
                                <td rowspan="3" class="right" style="font-size:11px;"><%=formatnumber(org_pay_count,0)%>&nbsp;명</td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_base_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_meals_pay,0)%></td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_postage_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_re_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_overtime_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_car_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_nps_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_nhis_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_epi_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_longcare_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_income_tax,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_wetax,0)%></td>
							</tr>
                            <tr>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_position_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_custom_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_job_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_job_support,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_jisa_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_long_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_disabled_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_other_amt1,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_sawo_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_school_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_nhis_bla_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_long_bla_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><strong><%=formatnumber(org_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;"><strong><%=formatnumber(org_give_tot,0)%></strong></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_hyubjo_amt,0)%></td>
                                
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_incom_tax,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_wetax,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_incom_tax2,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_wetax2,0)%></td>
                                <td align="right" style="font-size:11px;"><strong><%=formatnumber(org_curr_pay,0)%></strong></td>
							</tr>           
           
	       <%    
		                    team_base_pay = 0 
				            team_meals_pay = 0
	                        team_postage_pay = 0
	                        team_re_pay = 0
	                        team_overtime_pay = 0
	                        team_car_pay = 0
	                        team_position_pay = 0
	                        team_custom_pay = 0
	                        team_job_pay = 0
	                        team_job_support = 0
              	            team_jisa_pay = 0
	                        team_long_pay = 0
	                        team_disabled_pay = 0
	                        team_family_pay = 0
	                        team_school_pay = 0
	                        team_qual_pay = 0
	                        team_other_pay1 = 0
	                        team_other_pay2 = 0
	                        team_other_pay3 = 0
	                        team_tax_yes = 0
	                        team_tax_no = 0
	                        team_tax_reduced = 0
	                        team_give_tot = 0
                            team_nps_amt = 0
                            team_nhis_amt = 0
                            team_epi_amt = 0
                            team_longcare_amt = 0
                            team_income_tax = 0
                            team_wetax = 0
	                        team_year_incom_tax = 0
                            team_year_wetax = 0
							team_year_incom_tax2 = 0
                            team_year_wetax2 = 0
                            team_other_amt1 = 0
                            team_sawo_amt = 0
                            team_hyubjo_amt = 0
                            team_school_amt = 0
                            team_nhis_bla_amt = 0
                            team_long_bla_amt = 0
	                        team_deduct_tot = 0
	
	                        team_pay_count = 0	
	                        team_curr_pay = 0
				 
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
	                        org_family_pay = 0
	                        org_school_pay = 0
	                        org_qual_pay = 0
	                        org_other_pay1 = 0
	                        org_other_pay2 = 0
	                        org_other_pay3 = 0
	                        org_tax_yes = 0
	                        org_tax_no = 0
	                        org_tax_reduced = 0
	                        org_give_tot = 0
                            org_nps_amt = 0
                            org_nhis_amt = 0
                            org_epi_amt = 0
                            org_longcare_amt = 0
                            org_income_tax = 0
                            org_wetax = 0
	                        org_year_incom_tax = 0
                            org_year_wetax = 0
							org_year_incom_tax2 = 0
                            org_year_wetax2 = 0
                            org_other_amt1 = 0
                            org_sawo_amt = 0
                            org_hyubjo_amt = 0
                            org_school_amt = 0
                            org_nhis_bla_amt = 0
                            org_long_bla_amt = 0
	                        org_deduct_tot = 0
	
	                        org_pay_count = 0	
	                        org_curr_pay = 0
				 
				            bi_org = pmg_saupbu
							bi_company = pmg_company
							bi_bonbu = pmg_bonbu
		              end if
		   
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
		  
		              org_pay_count = org_pay_count + 1
		              org_base_pay = org_base_pay + int(rs("pmg_base_pay"))
	                  org_meals_pay = org_meals_pay + int(rs("pmg_meals_pay"))
	                  org_postage_pay = org_postage_pay + int(rs("pmg_postage_pay"))
	                  org_re_pay = org_re_pay + int(rs("pmg_re_pay"))
	                  org_overtime_pay = org_overtime_pay + int(rs("pmg_overtime_pay"))
	                  org_car_pay = org_car_pay + int(rs("pmg_car_pay"))
                      org_position_pay = org_position_pay + int(rs("pmg_position_pay"))
	                  org_custom_pay = org_custom_pay + int(rs("pmg_custom_pay"))
	                  org_job_pay = org_job_pay + int(rs("pmg_job_pay"))
	                  org_job_support = org_job_support + int(rs("pmg_job_support"))
	                  org_jisa_pay = org_jisa_pay + int(rs("pmg_jisa_pay"))
	                  org_long_pay = org_long_pay + int(rs("pmg_long_pay"))
	                  org_disabled_pay = org_disabled_pay + int(rs("pmg_disabled_pay"))
	                  org_give_tot = org_give_tot + int(rs("pmg_give_total"))
		  
		              team_pay_count = team_pay_count + 1
		              team_base_pay = team_base_pay + int(rs("pmg_base_pay"))
	                  team_meals_pay = team_meals_pay + int(rs("pmg_meals_pay"))
	                  team_postage_pay = team_postage_pay + int(rs("pmg_postage_pay"))
	                  team_re_pay = team_re_pay + int(rs("pmg_re_pay"))
	                  team_overtime_pay = team_overtime_pay + int(rs("pmg_overtime_pay"))
	                  team_car_pay = team_car_pay + int(rs("pmg_car_pay"))
                      team_position_pay = team_position_pay + int(rs("pmg_position_pay"))
	                  team_custom_pay = team_custom_pay + int(rs("pmg_custom_pay"))
	                  team_job_pay = team_job_pay + int(rs("pmg_job_pay"))
	                  team_job_support = team_job_support + int(rs("pmg_job_support"))
	                  team_jisa_pay = team_jisa_pay + int(rs("pmg_jisa_pay"))
	                  team_long_pay = team_long_pay + int(rs("pmg_long_pay"))
	                  team_disabled_pay = team_disabled_pay + int(rs("pmg_disabled_pay"))
	                  team_give_tot = team_give_tot + int(rs("pmg_give_total"))

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
		   
	            	  org_nps_amt = org_nps_amt + de_nps_amt
                      org_nhis_amt = org_nhis_amt + de_nhis_amt
                      org_epi_amt = org_epi_amt + de_epi_amt
		              org_longcare_amt = org_longcare_amt + de_longcare_amt
                      org_income_tax = org_income_tax + de_income_tax
                      org_wetax = org_wetax + de_wetax
		              org_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
                      org_year_wetax = sum_year_wetax + de_year_wetax
					  org_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
                      org_year_wetax2 = sum_year_wetax2 + de_year_wetax2
                      org_other_amt1 = org_other_amt1 + de_other_amt1
                      org_sawo_amt = org_sawo_amt + de_sawo_amt
                      org_hyubjo_amt = org_hyubjo_amt + de_hyubjo_amt
                      org_school_amt = org_school_amt + de_school_amt
                      org_nhis_bla_amt = org_nhis_bla_amt + de_nhis_bla_amt
                      org_long_bla_amt = org_long_bla_amt + de_long_bla_amt
		              org_deduct_tot = org_deduct_tot + de_deduct_tot
		   
		              team_nps_amt = team_nps_amt + de_nps_amt
                      team_nhis_amt = team_nhis_amt + de_nhis_amt
                      team_epi_amt = team_epi_amt + de_epi_amt
		              team_longcare_amt = team_longcare_amt + de_longcare_amt
                      team_income_tax = team_income_tax + de_income_tax
                      team_wetax = team_wetax + de_wetax
		              team_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
                      team_year_wetax = sum_year_wetax + de_year_wetax
					  team_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
                      team_year_wetax2 = sum_year_wetax2 + de_year_wetax2
                      team_other_amt1 = team_other_amt1 + de_other_amt1
                      team_sawo_amt = team_sawo_amt + de_sawo_amt
                      team_hyubjo_amt = team_hyubjo_amt + de_hyubjo_amt
                      team_school_amt = team_school_amt + de_school_amt
                      team_nhis_bla_amt = team_nhis_bla_amt + de_nhis_bla_amt
                      team_long_bla_amt = team_long_bla_amt + de_long_bla_amt
		              team_deduct_tot = team_deduct_tot + de_deduct_tot

  				      rs.movenext()
				  loop
				  rs.close()
						
				  sum_curr_pay = sum_give_tot - sum_deduct_tot
				  team_curr_pay = team_give_tot - team_deduct_tot
	              org_curr_pay = org_give_tot - org_deduct_tot
				            
							if bi_org = "" and bi_bonbu = "" then
							        bi_org = bi_company 
							   else
							       if bi_org = "" then 
								       bi_org = bi_bonbu + " 직할"
								   end if
						    end if
						
	       %>
                              <tr>
                                <td rowspan="3" class="first"><%=bi_org%>&nbsp;&nbsp;계</td>
                                <td rowspan="3" class="right" style="font-size:11px;"><%=formatnumber(org_pay_count,0)%>&nbsp;명</td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_base_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_meals_pay,0)%></td>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_postage_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_re_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_overtime_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_car_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_nps_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_nhis_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_epi_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_longcare_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_income_tax,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_wetax,0)%></td>
							</tr>
                            <tr>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_position_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_custom_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_job_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_job_support,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_jisa_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_long_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_disabled_pay,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_other_amt1,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_sawo_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_school_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_nhis_bla_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_long_bla_amt,0)%></td>
                                <td align="right" style="font-size:11px;"><strong><%=formatnumber(org_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
                                <td align="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td align="right" style="font-size:11px;"><strong><%=formatnumber(org_give_tot,0)%></strong></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_hyubjo_amt,0)%></td>
                                
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_incom_tax,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_wetax,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_incom_tax2,0)%></td>
                                <td align="right" style="font-size:11px;"><%=formatnumber(org_year_wetax2,0)%></td>
                                <td align="right" style="font-size:11px;"><strong><%=formatnumber(org_curr_pay,0)%></strong></td>
							</tr>

                          	<tr>
                                <td rowspan="3" class="first" style="background:#ffe8e8;">총계</td>
                                <td rowspan="3" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(pay_count,0)%>&nbsp;명</td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_base_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_meals_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_postage_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_re_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_overtime_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_car_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nps_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_epi_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_longcare_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_income_tax,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_wetax,0)%></td>
							</tr>
                            <tr>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_position_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_custom_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_support,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_jisa_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_disabled_pay,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_other_amt1,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_sawo_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_school_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_bla_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_bla_amt,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_give_tot,0)%></strong></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_hyubjo_amt,0)%></td>
                                
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax2,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax2,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_curr_pay,0)%></strong></td>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

