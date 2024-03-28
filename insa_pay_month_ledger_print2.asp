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

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
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

	rs.movenext()
loop
rs.close()

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_emp_no ASC"

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 개인별 급여대장(" + view_condi + ")"

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
				<form action="insa_pay_month_ledger_print2.asp" method="post" name="frm">
				<div class="gView">
                <table width="1150" cellpadding="0" cellspacing="0">
                   <tr>
                      <td class="style20C"><%=title_line%></td>
                   </tr>
                   <tr>
                      <td height="20" class="style20C">&nbsp;</td>
                   </tr>
                </table>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
							<col width="6%" >
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
				               <th colspan="2" height="30" scope="col" style=" border-bottom:1px solid #e3e3e3;">인적사항</th>
				               <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;">공제 및 차인지급액</th>
			                </tr>
                            <tr>
								<td height="30" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">사번</td> 
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">성  명</td>
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
								<td height="30" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">입사일</td> 
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">직급</td>
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
								<td height="30" scope="col" style=" border-bottom:2px solid #515254; background:#f8f8f8;">퇴사일</td> 
								<td scope="col" style=" border-bottom:2px solid #515254; background:#f8f8f8;">부서</td>
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
							  emp_no = rs("pmg_emp_no")
							  pmg_give_tot = rs("pmg_give_total")

	           			%>
							<tr>
								<td class="first"><%=rs("pmg_emp_no")%>&nbsp;</td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_emp_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_meals_pay"),0)%></td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_postage_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_re_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_car_pay"),0)%></td>
                        <%
						      Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not rs_emp.eof then
									emp_first_date = rs_emp("emp_first_date")
									emp_in_date = rs_emp("emp_in_date")
									emp_end_date = rs_emp("emp_end_date")
									emp_company = rs_emp("emp_company")
									emp_bonbu = rs_emp("emp_bonbu")
									emp_saupbu = rs_emp("emp_saupbu")
									emp_team = rs_emp("emp_team")
	                             else
									emp_first_date = ""
									emp_in_date = ""
									emp_company = ""
									emp_bonbu = ""
									emp_saupbu = ""
									emp_team = ""
                              end if
                              rs_emp.close()
							  if emp_end_date = "1999-01-01" then emp_end_date = "" end if
                          %>

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
						  
                          %>
                                <td class="right"><%=formatnumber(de_nps_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_nhis_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_epi_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_longcare_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(de_wetax,0)%></td>
							</tr>
                            <tr>
								<td class="first"><%=emp_in_date%>&nbsp;</td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_grade")%></td>
                                <td class="right"><%=formatnumber(rs("pmg_position_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_custom_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_job_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_job_support"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_jisa_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_long_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_disabled_pay"),0)%></td>
                                <td class="right"><%=formatnumber(de_other_amt1,0)%></td>
                                <td class="right"><%=formatnumber(de_sawo_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_school_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_nhis_bla_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_long_bla_amt,0)%></td>
                                <td class="right"><strong><%=formatnumber(de_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
								<td class="first"><%=emp_end_date%>&nbsp;</td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_org_name")%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><strong><%=formatnumber(rs("pmg_give_total"),0)%></strong></td>
                                <td class="right"><%=formatnumber(de_hyubjo_amt,0)%></td>
                                
                                <td class="right"><%=formatnumber(de_year_incom_tax,0)%></td>
                                <td class="right"><%=formatnumber(de_year_wetax,0)%></td>
                                <td class="right"><%=formatnumber(de_year_incom_tax2,0)%></td>
                                <td class="right"><%=formatnumber(de_year_wetax2,0)%></td>
                                <td class="right"><strong><%=formatnumber(pmg_curr_pay,0)%></strong></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot - sum_deduct_tot
						
						%>
                          	<tr>
                                <td rowspan="3" class="first" style="background:#ffe8e8;">총계</td>
                                <td rowspan="3" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(pay_count,0)%>&nbsp;명</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_base_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_meals_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_postage_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_re_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_overtime_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sumpmg_car_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nps_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_epi_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_longcare_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_income_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_wetax,0)%></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_position_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_custom_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_support,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_jisa_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_disabled_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_other_amt1,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_sawo_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_school_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_bla_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_bla_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3; font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_give_tot,0)%></strong></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_hyubjo_amt,0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax2,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax2,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_curr_pay,0)%></strong></td>
							</tr>
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

