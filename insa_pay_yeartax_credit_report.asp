<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

'on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

inc_yyyy = cint(mid(now(),1,4)) - 1

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_cred = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")
emp_person = cstr(rs_emp("emp_person1")) + "-" + cstr(rs_emp("emp_person2"))	
rs_emp.close()	

tot_pay = 0
y_nhis_amt = 0
y_longcare_amt = 0
y_epi_amt = 0
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
if not rs_year.eof then
       y_nhis_amt = rs_year("y_nhis_amt")
	   y_longcare_amt = rs_year("y_longcare_amt")
	   y_epi_amt = rs_year("y_epi_amt")
	   tot_pay = rs_year("y_total_pay") + rs_year("y_total_bonus") + rs_year("y_other_pay")
   else
       y_nhis_amt = 0
	   y_longcare_amt = 0
	   y_epi_amt = 0
end if

Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
'Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
	   tot_pay = tot_pay + rs_bef("b_pay") + rs_bef("b_bonus") + rs_bef("b_deem_bonus")
	rs_bef.MoveNext()
loop
rs_bef.close()

if emp_company = "케이원정보통신" then
      company_name = "(주)" + "케이원정보통신"
	  owner_name = "김승일"
	  addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif emp_company = "휴디스" then
              company_name = "(주)" + "휴디스"
			  owner_name = "김한종"
	          addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif emp_company = "케이네트웍스" then
                     company_name = "케이네트웍스" + "(주)"
					 owner_name = "이중원"
	                 addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif emp_company = "에스유에이치" then
                        company_name = "(주)" + "에스유에이치"	
						owner_name = "박미애"
	                    addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

c_hap1 = 0

'sql = "select * from pay_yeartax_credit where c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' ORDER BY c_emp_no,c_person_no,c_id,c_seq ASC"
'Rs.Open Sql, Dbconn, 1
sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산-신용카드등 소득공제 신청서"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.inc_yyyy.value == "") {
					alert ("귀속년도를 입력하세요.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_credit_report.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="20%" >
							<col width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td colspan="4">소득자 인적 사항</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-top:1px solid #e3e3e3;">성명</th>
                              <td><%=emp_name%></td>
                              <th class="left" style=" border-top:1px solid #e3e3e3;">주민등록번호(또는 외국인등록번호)</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th class="left">법인명</th>
                              <td><%=company_name%></td>
                              <th class="left">업체명</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
                              <td colspan="4">&nbsp;</td>
						    </tr>
						</thead>
					</table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="5%" >
                              <col width="5%" >
                              <col width="*" >
                              <col width="10%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="10%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1. 공제대상자 및 공제대상금액 명세</th>
                              </tr>
                              <tr>
                                <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">공제대상자</th>
                                <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">신용카드등 사용금액</th>
                              </tr>
                              <tr>
                                <th class="first" scope="col">1내·외국인구분</th>
                                <th scope="col">2관계</th>
                                <th scope="col">3성명</th>
                                <th scope="col">4생년월일</th>
                                <th scope="col">자료구분</th>
                                <th scope="col">5소계<br>(6+7+8+9+10)</th>
                                <th scope="col">6신용카드<br>(전통시장·대중교통 제외)</th>
                                <th scope="col">7현금영수중<br>(전통시장·대중교통 제외)</th>
                                <th scope="col">8직불·선불카드<br>(전통시장·대중교통 제외)</th>
                                <th scope="col">9전통시장사용분<br>(신용카드,직불·선불카드,현금영수증)</th>
                                <th scope="col">10대중교통<br>(신용카드,직불·선불카드,현금영수증)</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						sum_market = 0
						sum_transit = 0
						sum_credit = 0
						sum_cash = 0
						sum_direct = 0
						sum_hap = 0
						
						do until rs.eof
						
						   chk_sum = Rs("c_credit_nts") + Rs("c_credit_other") + Rs("c_cash_nts") + Rs("c_direct_nts") + Rs("c_direct_other") + Rs("c_market_nts") + Rs("c_market_other") + Rs("c_transit_nts") + Rs("c_transit_other")
						
                          if chk_sum > 0 then 
                             f_national = Rs("f_national") 						 

							 nts_market = Rs("c_market_nts")
							 nts_transit = Rs("c_transit_nts")
							 other_market = Rs("c_market_other")
							 other_transit = Rs("c_transit_other")
							 nts_hap = Rs("c_credit_nts") + Rs("c_cash_nts") + Rs("c_direct_nts") + nts_market + nts_transit
							 other_hap = Rs("c_credit_other") + Rs("c_direct_other") + other_market + other_transit
							 
							 sum_market = sum_market + nts_market + other_market
							 sum_transit = sum_transit + nts_transit + other_transit
							 sum_credit = sum_credit + Rs("c_credit_nts") + Rs("c_credit_other")
							 sum_cash = sum_cash + Rs("c_cash_nts")
							 sum_direct =  sum_direct + Rs("c_direct_nts") + Rs("c_direct_other")
							 sum_hap =  sum_hap + nts_hap + other_hap
	           			%>
							<tr>
                                <td rowspan="2"><%=f_national%>&nbsp;</td>
                                <td rowspan="2"><%=Rs("f_rel")%>&nbsp;</td>
                                <td rowspan="2"><%=Rs("f_family_name")%>&nbsp;</td>
                                <td rowspan="2"><%=Rs("f_person_no")%>&nbsp;</td>
                                <td class="left">국세청 자료</td>
                                <td class="right"><%=formatnumber(nts_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("c_credit_nts"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("c_cash_nts"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("c_direct_nts"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(nts_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(nts_transit,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td class="left" style=" border-left:1px solid #e3e3e3;">그 밖의 자료</td>
                                <td class="right"><%=formatnumber(other_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("c_credit_other"),0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("c_direct_other"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(other_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(other_transit,0)%>&nbsp;</td>
							</tr>
						<%
						    end if
							rs.movenext()
						loop
						rs.close()
						    if sum_hap > 0 then
						%>
                        	<tr>
                                <td colspan="5">10합계액</td>
                                <td class="right"><%=formatnumber(sum_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_credit,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_cash,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_direct,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_transit,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="11">&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">2. 신용카드등 소득공제액의 계산</th>
                            </tr>
                        <%
						    tax15_31 = 0
							tax15_32 = 0
							market_tax = int(sum_market * (30 / 100))
							transit_tax = int(sum_transit * (30 / 100))
							cash_tax = int((sum_cash + sum_direct) * (30 / 100))
							credit_tax = int(sum_credit * (15 / 100))
							pay_tax = int(tot_pay * (25 / 100))
							if pay_tax <= sum_credit then 
							       tax15_31 = int(pay_tax * (15 / 100))
								   tax15_3 = tax15_31
							   else
							       tax15_32 = int((sum_credit * (15 / 100)) + ((pay_tax - sum_credit) * (30 / 100)))
								   tax15_3 = tax15_32
							end if
						%>
                            <tr>
                                <th rowspan="2" colspan="3" style="background:#f8f8f8;">11전통시장사용분<br>공제액<br>(9*30%)</th>
                                <th rowspan="2" colspan="2" style="background:#f8f8f8;">12대중교통이용분<br>공제액<br>(10*30%)</th>
                                <th rowspan="2" style="background:#f8f8f8;">13직불·선불카드<br>현금영수증등 사용분<br>(7+8)*30%</th>
                                <th rowspan="2" style="background:#f8f8f8;">14신용카드사용분<br>공제액<br>(6*15%)</th>
                                <th colspan="3" style="background:#f8f8f8;">15공제제외금액 계산</th>
                                <th rowspan="2" style="background:#f8f8f8;">16체크카드등<br>사용액증가분<br>공제액</th>
							</tr>
                            <tr>
                                <th style="background:#f8f8f8; border-left:1px solid #e3e3e3;"">15-1<br>총급여</th>
                                <th style="background:#f8f8f8;">15-2<br>최저사용금액<br>(15-1*25%)</th>
                                <th style="background:#f8f8f8;">15-3<br>공제제외금액</th>
							</tr>
                            <tr>
                                <td colspan="3" class="right"><%=formatnumber(market_tax,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(transit_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cash_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(credit_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(tot_pay,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pay_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(tax15_3,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="3" style="background:#f8f8f8;">17공제가능금액<br>[11+12+13+14-(15-3)+16]<br>(9*30%)</th>
                                <th colspan="2" style="background:#f8f8f8;">18공제한도액<br>[3백만원과<br>(15-1)*20%중 적은금액]</th>
                                <th style="background:#f8f8f8;">19일반 공제금액<br>(17과 18중 작은금액)</th>
                                <th colspan="2" style="background:#f8f8f8;">20전통시장 추가공제금액<br>[17-18(음수이면 0으로 봄)과<br>11중 적은금액(한도:1백만원)]</th>
                                <th colspan="2" style="background:#f8f8f8;">21대중교통 추가공제금액<br>[17-20-19(음수이면 0으로 봄)과<br>12중 적은금액(한도:1백만원)]</th>
                                <th style="background:#f8f8f8;">22최종 공제금액<br>[19+20+21]</th>
							</tr>
                            <tr>
                                <td colspan="3" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">15-3 계산</th>
                            </tr>
                            <tr>
                                <th colspan="4" style="background:#f8f8f8;">구분</th>
                                <th colspan="5" style="background:#f8f8f8;">계산식</th>
                                <th colspan="2" style="background:#f8f8f8;">15-3</th>
							</tr>
                            <tr>
                                <th colspan="4" style="background:#f8f8f8;">15-2 최저사용금액 ≤ 신용카드사용분6</th>
                                <td colspan="5" class="left" >15-2 * 15%</td>
                                <td colspan="2" class="right"><%=formatnumber(tax15_31,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="4" style="background:#f8f8f8;">15-2 최저사용금액 > 신용카드사용분6</th>
                                <td colspan="5" class="left" >6 * 15% + [(15-2) - 6] * 30%</td>
                                <td colspan="2" class="right"><%=formatnumber(tax15_32,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">16 계산</th>
                            </tr>
                            <tr>
                                <th colspan="3" style="background:#f8f8f8;">구분</th>
                                <th style="background:#f8f8f8;">과세기간</th>
                                <th colspan="2" style="background:#f8f8f8;">금액</th>
                                <th colspan="5" class="left" style="background:#f8f8f8;">16체크카드 등 사용액 증가분 공제액</th>
							</tr>
                            <tr>
                                <th rowspan="2" colspan="3" style="background:#f8f8f8;">본인의 신용카드등 사용액</th>
                                <th style="background:#f8f8f8;">2013년</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td rowspan="2" colspan="5" class="left" >&nbsp;</td>
							</tr>
                            <tr>
                                <th style=" border-left:1px solid #e3e3e3; background:#f8f8f8;">2014년</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th rowspan="2" colspan="3" style="background:#f8f8f8;">본인의 신용카드등 사용액</th>
                                <th style="background:#f8f8f8;">2013년</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td rowspan="2" colspan="5" class="left" >(계산식)<br>㉮2013년 본인의 신용카드 등 사용액 ≥ 2014년 본인의 신용카드 등 사용액 : "0"<br>㉯2013년 본인의 신용카드 등 사용액 < 2014년 본인의 신용카드등<br>사용액:(2014년 하반기 추가공제율사용분 - 2013년 추가공제율사용분<br>*50%) * 10%(단, 음수인 경우 "0")</td>
							</tr>
                            <tr>
                                <th style="border-left:1px solid #e3e3e3; background:#f8f8f8;">2014년<br>하반기</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="11">&nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">「조세특례제한법 시행령」제121조의2제8항에 따라 신용카드 등 사용금액에 대한 소득공제를 신청 합니다.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 년 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;월 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;일<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;신청인 : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(서명 또는 인)<br></td>
                            </tr>
                            <tr>
                                <td colspan="11">&nbsp;</td>
                            </tr>
                            <tr>
                                <th style="background:#f8f8f8;">구비서류</th>
                                <td colspan="9" class="left" >신용카드 등 사용금액 확인서(별지 제74호의5서식을 말합니다) 또는 국세청 홈페이지에서 제공하는 신용카드 등 사용금액 명세를 출력한 서류 1부</td>
                                <td>수수료 없음</td>
							</tr>
                   <%
				          end if
				   %>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType04">의료비지급명세서</a>
                    <a href="insa_pay_yeartax_donation_report.asp" class="btnType04">기부금명세서</a>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_credit_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_credit_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">신용카드등 명세서 출력</a>
                    <a href="insa_pay_yeartax_tax_report.asp" class="btnType04">소득공제신고서</a>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

