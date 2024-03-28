<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

view_condi=Request("view_condi")
from_yymm=request("from_yymm")
to_yymm=request("to_yymm")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)
	
f_yymm = from_yymm
t_yymm = to_yymm

title_line = cstr(f_yymm) + " ∼ " + cstr(t_yymm) + "월 " + " 급여항목별 집계-" + view_condi 

savefilename = cstr(f_yymm) + " ∼ " + cstr(t_yymm) + "월 급여항목별 집계.xls"

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
	
	tax_meals_no = 0	
	tax_car_no = 0	
	tax_meals_yes = 0	
	tax_car_yes = 0	
	
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
          Sql = "select * from pay_month_give where (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1')"
	else	  
		  Sql = "select * from pay_month_give where (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
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
	  
	  'sum_tax_yes = sum_tax_yes + int(rs("pmg_tax_yes"))
	  'sum_tax_no = sum_tax_no + int(rs("pmg_tax_no"))
	  'sum_tax_reduced = sum_tax_reduced + int(rs("pmg_tax_reduced"))
	  
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

	  meals_pay = pmg_meals_pay
	  car_pay = pmg_car_pay
	  meals_tax_pay = 0
	  meals_taxno_pay = 0
	  car_tax_pay = 0
	  car_taxno_pay = 0
	  
	  if  meals_pay > 100000 then
	         meals_tax_pay = meals_pay - 100000
	         tax_meals_yes = tax_meals_yes + (meals_pay - 100000)
			 meals_taxno_pay = 100000
			 tax_meals_no= tax_meals_no + 100000
		  else	 
		     meals_taxno_pay = meals_pay
			 tax_meals_no= tax_meals_no + meals_pay
	  end if
  	  if car_pay > 200000 then
	         car_tax_pay = car_pay - 200000
			 tax_car_yes = tax_car_yes + (car_pay - 200000)
			 car_taxno_pay = 200000
			 tax_car_no =  tax_car_no + 200000
		 else
			 tax_car_no =  tax_car_no + car_pay
			 car_taxno_pay = car_pay
	  end if
	  
	  pmg_tax_yes = 0
	  pmg_tax_no = 0
	  
	  pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	  pmg_tax_no = meals_taxno_pay + car_taxno_pay
	  
	  sum_tax_yes = sum_tax_yes + pmg_tax_yes
	  sum_tax_no = sum_tax_no + pmg_tax_no
	  
	rs.movenext()
loop
rs.close()		

if view_condi = "전체" then
          Sql = "select * from pay_month_deduct where (de_yymm >= '"+f_yymm+"' and de_yymm <= '"+t_yymm+"') and (de_id = '1')"
	else	  
		  Sql = "select * from pay_month_deduct where (de_yymm >= '"+f_yymm+"' and de_yymm <= '"+t_yymm+"') and (de_id = '1') and (de_company = '"+view_condi+"')"
end if					  
Set Rs_dct = DbConn.Execute(SQL)							  
							  
do until Rs_dct.eof
	  sum_nps_amt = sum_nps_amt + int(Rs_dct("de_nps_amt"))
      sum_nhis_amt = sum_nhis_amt + int(Rs_dct("de_nhis_amt"))
      sum_epi_amt = sum_epi_amt + int(Rs_dct("de_epi_amt"))
      sum_longcare_amt = sum_longcare_amt + int(Rs_dct("de_longcare_amt"))
      sum_income_tax = sum_income_tax + int(Rs_dct("de_income_tax"))
      sum_wetax = sum_wetax + int(Rs_dct("de_wetax"))
	  sum_year_incom_tax = sum_year_incom_tax + int(Rs_dct("de_year_incom_tax"))
      sum_year_wetax = sum_year_wetax + int(Rs_dct("de_year_wetax"))
	  sum_year_incom_tax2 = sum_year_incom_tax2 + int(Rs_dct("de_year_incom_tax2"))
      sum_year_wetax2 = sum_year_wetax2 + int(Rs_dct("de_year_wetax2"))
      sum_other_amt1 = sum_other_amt1 + int(Rs_dct("de_other_amt1"))
      sum_sawo_amt = sum_sawo_amt + int(Rs_dct("de_sawo_amt"))
      sum_hyubjo_amt = sum_hyubjo_amt + int(Rs_dct("de_hyubjo_amt"))
      sum_school_amt = sum_school_amt + int(Rs_dct("de_school_amt"))
      sum_nhis_bla_amt = sum_nhis_bla_amt + int(Rs_dct("de_nhis_bla_amt"))
      sum_long_bla_amt = sum_long_bla_amt + int(Rs_dct("de_long_bla_amt"))	
      sum_deduct_tot = sum_deduct_tot + int(Rs_dct("de_deduct_total"))
	Rs_dct.movenext()
loop
Rs_dct.close()		

sum_curr_pay = sum_give_tot - sum_deduct_tot

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th colspan="4" class="first" style="background:#F5FFFA">지&nbsp;급&nbsp;&nbsp;&nbsp;항&nbsp;목</th>
								<th colspan="4" class="first" style="background:#F8F8FF">공&nbsp;제&nbsp;&nbsp;&nbsp;항&nbsp;목</th>
							</tr>  
                        </thead>
                        <tbody>
							<tr>
								<th class="first" style="background:#F5FFFA">기본급</th>
								<td align="right"><%=formatnumber(sum_base_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">식대(비과세)</th>
								<td align="right"><%=formatnumber(tax_meals_no,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">국민연금</th>
                                <td align="right"><%=formatnumber(sum_nps_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">건강보험</th>
                                <td align="right"><%=formatnumber(sum_nhis_amt,0)%>&nbsp;</td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F5FFFA">통신비</th>
								<td align="right"><%=formatnumber(sum_postage_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">식대</th>
								<td align="right"><%=formatnumber(tax_meals_yes,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">고용보험</th>
                                <td align="right"><%=formatnumber(sum_epi_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">장기요양보험</th>
                                <td align="right"><%=formatnumber(sum_longcare_amt,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#F5FFFA">연장근로수당</th>
								<td align="right"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">소급급여</th>
								<td align="right"><%=formatnumber(sum_re_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">소득세</th>
                                <td align="right"><%=formatnumber(sum_income_tax,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">지방소득세</th>
                                <td align="right"><%=formatnumber(sum_wetax,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#F5FFFA">직책수당</th>
								<td align="right"><%=formatnumber(sum_position_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">주차지원금(비과세)</th>
								<td align="right"><%=formatnumber(tax_car_no,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">기타공제</th>
                                <td align="right"><%=formatnumber(sum_other_amt1,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">경조회비</th>
                                <td align="right"><%=formatnumber(sum_sawo_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">직무보조비</th>
								<td align="right"><%=formatnumber(sum_job_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">주차지원금</th>
								<td align="right"><%=formatnumber(tax_car_yes,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">협조비</th>
                                <td align="right"><%=formatnumber(sum_hyubjo_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">학자금대출</th>
                                <td align="right"><%=formatnumber(sum_school_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">본지사근무비</th>
								<td align="right"><%=formatnumber(sum_jisa_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">고객관리수당</th>
								<td align="right"><%=formatnumber(sum_custom_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">건강보험료정산</th>
                                <td align="right"><%=formatnumber(sum_nhis_bla_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">장기요양보험정산</th>
                                <td align="right"><%=formatnumber(sum_long_bla_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">근속수당</th>
								<td align="right"><%=formatnumber(sum_long_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">업무장려비</th>
								<td align="right"><%=formatnumber(sum_job_support,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">연말정산소득세</th>
                                <td align="right"><%=formatnumber(sum_year_incom_tax,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">연말정산지방세</th>
                                <td align="right"><%=formatnumber(sum_year_wetax,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style=" border-bottom:2px solid #515254; background:#F5FFFA">장애인수당</th>
								<td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_disabled_pay,0)%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td align="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">연말재정산소득세</th>
                                <td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_year_incom_tax2,0)%>&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">연말재정산지방세</th>
                                <td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_year_wetax2,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">과세</th>
								<td align="right"><%=formatnumber(sum_tax_yes,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">&nbsp;</th>
								<td align="right">&nbsp;</td>
								<th style="background:#F8F8FF">&nbsp;</th>
                                <td align="left">&nbsp;</td>
                                <th style="background:#F8F8FF">&nbsp;</th>
                                <td align="right">&nbsp;</td>
							</tr>      
                            <tr>
								<th class="first" style="background:#F5FFFA">비과세</th>
								<td align="right"><%=formatnumber(sum_tax_no,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">&nbsp;</th>
								<td align="right">&nbsp;</td>
								<th style="background:#F8F8FF">&nbsp;</th>
                                <td align="left">&nbsp;</td>
                                <th style="background:#F8F8FF">&nbsp;</th>
                                <td align="right">&nbsp;</td>
							</tr>  
                            <tr>
								<th class="first" style="border-bottom:2px solid #515254; background:#F5FFFA">감면소득</th>
								<td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_tax_reduced,0)%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td align="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">&nbsp;</th>
                                <td align="left" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F8F8FF">&nbsp;</th>
                                <td align="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="border-bottom:2px solid #515254; background:#F5FFFA">지급액 계</th>
								<td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_give_tot,0)%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td align="right" style=" border-bottom:2px solid #515254;"><%=pay_count%>&nbsp;명</td>
                                <th style="border-bottom:2px solid #515254; background:#F8F8FF">공제액 계</th>
                                <td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_deduct_tot,0)%>&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">차인지급액</th>
                                <td align="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</td>
							</tr>              
						</tbody>
					</table>
		</div>	
      </div>			
	</div> 	
  </body>
</html>

