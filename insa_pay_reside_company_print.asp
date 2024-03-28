<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

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
dim sum_year_incom_tax2(6)
dim sum_year_wetax2(6)
dim sum_other_amt1(6)
dim sum_sawo_amt(6)
dim sum_hyubjo_amt(6)
dim sum_school_amt(6)
dim sum_nhis_bla_amt(6)
dim sum_long_bla_amt(6)
dim sum_deduct_tot(6)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

	
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
		sum_year_incom_tax2(i) = 0
        sum_year_wetax2(i) = 0
        sum_other_amt1(i) = 0
        sum_sawo_amt(i) = 0
        sum_hyubjo_amt(i) = 0
        sum_school_amt(i) = 0
        sum_nhis_bla_amt(i) = 0
        sum_long_bla_amt(i) = 0
        sum_deduct_tot(i) = 0
    next

	sum_curr_pay = 0	
	
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

order_Sql = " ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
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
	pmg_yymm = rs("pmg_yymm")
	
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
				 sum_year_incom_tax2(i) = sum_year_incom_tax2(i) + de_year_incom_tax2
                 sum_year_wetax2(i) = sum_year_wetax2(i) + de_year_wetax2
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
                 sum_wetax(6) = sum_wetax(5) + de_wetax
	             sum_year_incom_tax(6) = sum_year_incom_tax(6) + de_year_incom_tax
                 sum_year_wetax(6) = sum_year_wetax(6) + de_year_wetax
				 sum_year_incom_tax2(6) = sum_year_incom_tax2(6) + de_year_incom_tax2
                 sum_year_wetax2(6) = sum_year_wetax2(6) + de_year_wetax2
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

if view_condi = "전체" then
      Sql = " SELECT a.cost_group, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select cost_group,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') group by cost_group " & _
			"   order by cost_group " & _
			"   ) a, " & _
			" ( select cost_group,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') group by cost_group " & _	
			"   order by cost_group " & _
			"   ) b " & _		
			"  WHERE a.cost_group = b.cost_group " & _
			"  ORDER BY a.cost_group ASC " 
    else
      Sql = " SELECT a.cost_group, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select cost_group,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') group by cost_group " & _
			"   order by cost_group " & _
			"   ) a, " & _
			" ( select cost_group,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_company = '"+view_condi+"') group by cost_group " & _	
			"   order by cost_group " & _
			"   ) b " & _		
			"  WHERE a.cost_group = b.cost_group " & _
			"  ORDER BY a.cost_group ASC " 
end if

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 상주회사별 급여현황(" + view_condi + ")"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
		  		 window.close () ;
			}
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }
        </script>
    <style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
    </style>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="wrap">			
			<div id="container">
				<form action="insa_pay_reside_company_print.asp" method="post" name="frm">
				<div class="gView">
                <table width="1150" cellpadding="0" cellspacing="0">
                   <tr>
                      <td class="style20C"><strong><%=title_line%></strong></td>
                   </tr>
                   <tr>
                      <td height="20" class="style20C">&nbsp;</td>
                   </tr>
                </table>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
							<col width="8%" >
							<col width="*" >
							<col width="8%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="7%" >
							<col width="8%" >
                            <col width="7%" >
                            <col width="6%" >
							<col width="6%" > 
                            <col width="6%" >
                            <col width="7%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
				               <th colspan="2" height="30" scope="col" style=" border-bottom:1px solid #e3e3e3;">상주회사&nbsp;명</th>
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
						do until rs.eof
							  pmg_give_tot = cdbl(rs("pmg_give_total"))
							  
							  pmg_curr_pay = cdbl(rs("pmg_give_total")) - cdbl(rs("de_deduct_total"))
							  
							  sub_give_hap = cdbl(rs("pmg_postage_pay")) + cdbl(rs("pmg_re_pay")) + cdbl(rs("pmg_car_pay")) + cdbl(rs("pmg_position_pay")) + cdbl(rs("pmg_custom_pay")) + cdbl(rs("pmg_job_pay")) + cdbl(rs("pmg_job_support")) + cdbl(rs("pmg_jisa_pay")) + cdbl(rs("pmg_long_pay")) + cdbl(rs("pmg_disabled_pay"))
							
							saupbu_name = rs("cost_group")
							if saupbu_name = "" or saupbu_name = " " or isnull(saupbu_name) then
							    saupbu_name = view_condi
							end if
							  
	           			%>
							<tr>
								<td rowspan="3" class="first"><%=saupbu_name%>&nbsp;</td>
                                <td rowspan="3" class="right" style="font-size:11px;"><%=rs("saup_count")%>&nbsp;명</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_base_pay"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_meals_pay"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_postage_pay"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_re_pay"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_overtime_pay"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_car_pay"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_nps_amt"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_nhis_amt"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_epi_amt"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_longcare_amt"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_income_tax"),0)%>&nbsp;</td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_wetax"),0)%>&nbsp;</td>
                       <%  
                                  pmg_curr_pay = cdbl(rs("pmg_give_total")) - cdbl(rs("de_deduct_total"))
							  
							      hap_de_insur = cdbl(rs("de_nps_amt")) + cdbl(rs("de_nhis_amt")) + cdbl(rs("de_epi_amt")) + cdbl(rs("de_longcare_amt"))
							      hap_de_tax = cdbl(rs("de_income_tax")) + cdbl(rs("de_wetax")) + cdbl(rs("de_year_incom_tax")) + cdbl(rs("de_year_wetax")) + cdbl(rs("de_year_incom_tax2")) + cdbl(rs("de_year_wetax2"))
							      hap_de_other = cdbl(rs("de_other_amt1")) + cdbl(rs("de_sawo_amt")) + cdbl(rs("de_hyubjo_amt")) + cdbl(rs("de_school_amt")) + cdbl(rs("de_nhis_bla_amt")) + cdbl(rs("de_long_bla_amt"))
								  hap_deduct_tot = hap_de_insur + hap_de_tax + hap_de_other
								  
                       %>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px;"><%=formatnumber(rs("pmg_position_pay"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_custom_pay"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_job_pay"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_job_support"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_jisa_pay"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_long_pay"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("pmg_disabled_pay"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_other_amt1"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_sawo_amt"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_school_amt"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_nhis_bla_amt"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_long_bla_amt"),0)%></td>
                                <td class="right" style="font-size:11px;"><strong><%=formatnumber(rs("de_deduct_total"),0)%></strong></td>
                            </tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;">&nbsp;</td>
                                <td class="right" style="font-size:11px;"><strong><%=formatnumber(rs("pmg_give_total"),0)%></strong></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_hyubjo_amt"),0)%></td>
                                
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_year_incom_tax"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_year_wetax"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_year_incom_tax2"),0)%></td>
                                <td class="right" style="font-size:11px;"><%=formatnumber(rs("de_year_wetax2"),0)%></td>
                                <td class="right" style="font-size:11px;"><strong><%=formatnumber(pmg_curr_pay,0)%></strong></td>
                            </tr>
					   <%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot(6) - sum_deduct_tot(6)
						
						sum_give_hap = sum_postage_pay(6) + sum_re_pay(6) + sum_car_pay(6) + sum_position_pay(6) + sum_custom_pay(6) + sum_job_pay(6) + sum_job_support(6) + sum_jisa_pay(6) + sum_long_pay(6) + sum_disabled_pay(6)
						sum_de_insur =sum_nps_amt(6) + sum_nhis_amt(6) + sum_epi_amt(6) + sum_longcare_amt(6)
						sum_de_tax =sum_income_tax(6) + sum_wetax(6) + sum_year_incom_tax(6) + sum_year_wetax(6) + sum_year_incom_tax2(6) + sum_year_wetax2(6)
						sum_de_other =sum_other_amt1(6) + sum_sawo_amt(6) + sum_hyubjo_amt(6) + sum_school_amt(6) + sum_nhis_bla_amt(6) + sum_long_bla_amt(6)

						%>
                          	<tr>
                                <td rowspan="3" class="first" style="background:#EEFFFF;">총계</td>
                                <td rowspan="3" class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(pay_count(6),0)%>&nbsp;명</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_base_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_meals_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_postage_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_re_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_overtime_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_car_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_nps_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_nhis_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_epi_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_longcare_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_income_tax(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_wetax(6),0)%></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_position_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_custom_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_job_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_job_support(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_jisa_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_long_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_disabled_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_other_amt1(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_sawo_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_school_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_nhis_bla_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_long_bla_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><strong><%=formatnumber(sum_deduct_tot(6),0)%></strong></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3; font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><strong><%=formatnumber(sum_give_tot(6),0)%></strong></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_hyubjo_amt(6),0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_year_incom_tax(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_year_wetax(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_year_incom_tax2(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><%=formatnumber(sum_year_wetax2(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#EEFFFF;"><strong><%=formatnumber(sum_curr_pay,0)%></strong></td>
							</tr>

                         <%
						    for i = 1 to 6 
                        	     if	com_tab(i) <> "" then
								 
								    sum_curr_pay = sum_give_tot(i) - sum_deduct_tot(i)
						
						            sum_give_hap = sum_postage_pay(i) + sum_re_pay(i) + sum_car_pay(i) + sum_position_pay(i) + sum_custom_pay(i) + sum_job_pay(i) + sum_job_support(i) + sum_jisa_pay(i) + sum_long_pay(i) + sum_disabled_pay(i)
						            sum_de_insur =sum_nps_amt(i) + sum_nhis_amt(i) + sum_epi_amt(i) + sum_longcare_amt(i)
						            sum_de_tax =sum_income_tax(i) + sum_wetax(i) + sum_year_incom_tax(i) + sum_year_wetax(i) + sum_year_incom_tax2(i) + sum_year_wetax2(i)
						            sum_de_other =sum_other_amt1(i) + sum_sawo_amt(i) + sum_hyubjo_amt(i) + sum_school_amt(i) + sum_nhis_bla_amt(i) + sum_long_bla_amt(i)
						 %>	
                            <tr>
                                <td rowspan="3" class="first" style="background:#ffe8e8;"><%=com_tab(i)%></td>
                                <td rowspan="3" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(pay_count(i),0)%>&nbsp;명</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_base_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_meals_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_postage_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_re_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_overtime_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_car_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nps_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_epi_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_longcare_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_income_tax(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_wetax(i),0)%></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_position_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_custom_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_support(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_jisa_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_disabled_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_other_amt1(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_sawo_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_school_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_bla_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_bla_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_deduct_tot(i),0)%></strong></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3; font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_give_tot(i),0)%></strong></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_hyubjo_amt(i),0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax2(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax2(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_curr_pay,0)%></strong></td>
							</tr>
                         <%
							     end if
						    next
					     %>
						</tbody>
					</table>
				</div>
				<table width="1150" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<br>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				    <br>                 
                    </td>
			      </tr>
				</table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

