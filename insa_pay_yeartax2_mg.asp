<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

y_final=Request("y_final")

be_pg = "insa_pay_yeartax2_mg.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

'inc_yyyy = mid(cstr(now()),1,4)
inc_yyyy = cint(mid(now(),1,4)) - 1
f_yymm = cstr(inc_yyyy) + "01"
t_yymm = cstr(inc_yyyy) + "12"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_fam = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Set Rs_bef = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_person1 = rs_emp("emp_person1")
emp_person2 = rs_emp("emp_person2")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")
y_disab = rs_emp("emp_disabled_yn")
emp_national = "대한민국"
rs_emp.close()	

y_householder = "N"
Y_foreign = ""
y_woman = ""
y_single = ""
y_blue = ""
y_live = "Y"
y_change = "N"

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

sum_give_tot = 0
sum_bunus_tot = 0
sum_other_tot = 0
sum_tax_no = 0
sum_nps_amt = 0
sum_nhis_amt = 0
sum_epi_amt = 0
sum_longcare_amt = 0
sum_income_tax = 0
sum_wetax = 0

Sql = "select * from pay_month_give where (pmg_yymm >= '"&f_yymm&"' and pmg_yymm <= '"&t_yymm&"') and (pmg_id = '1') and (pmg_emp_no = '"&emp_no&"') and (pmg_company = '"&emp_company&"')"
Rs_give.Open Sql, Dbconn, 1
'Set Rs_give = DbConn.Execute(SQL)
do until Rs_give.eof
       pmg_yymm = Rs_give("pmg_yymm")
	   pay_year = mid(cstr(Rs_give("pmg_yymm")),1,4)
            pmg_give_tot = int(Rs_give("pmg_give_total"))	
		    meals_pay = int(Rs_give("pmg_meals_pay"))	
			car_pay = int(Rs_give("pmg_car_pay"))	
	        if  meals_pay > 100000 then
			    meals_pay =  100000
	        end if
	        if  car_pay > 200000 then
			    car_pay =  200000
	        end if
	        sum_give_tot = sum_give_tot + pmg_give_tot
	        sum_tax_no = sum_tax_no + meals_pay + car_pay

  		    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"&emp_no&"') and (de_company = '"&emp_company&"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_nps_amt = int(Rs_dct("de_nps_amt"))	
					 de_nhis_amt = int(Rs_dct("de_nhis_amt"))	
					 de_epi_amt = int(Rs_dct("de_epi_amt"))	
					 de_longcare_amt = int(Rs_dct("de_longcare_amt"))	
					 de_income_tax = int(Rs_dct("de_income_tax"))	
					 de_wetax = int(Rs_dct("de_wetax"))	
                  else
                     de_nps_amt = 0
					 de_nhis_amt = 0
					 de_epi_amt = 0
					 de_longcare_amt = 0
					 de_income_tax = 0
					 de_wetax = 0
              end if
              Rs_dct.close()
			  sum_nps_amt = sum_nps_amt + de_nps_amt
	          sum_nhis_amt = sum_nhis_amt + de_nhis_amt
			  sum_epi_amt = sum_epi_amt + de_epi_amt
	          sum_longcare_amt = sum_longcare_amt + de_longcare_amt
			  sum_income_tax = sum_income_tax + de_income_tax
	          sum_wetax = sum_wetax + de_wetax
	Rs_give.MoveNext()
loop
Rs_give.close()
'상여금
Sql = "select * from pay_month_give where (pmg_yymm >= '"&f_yymm&"' and pmg_yymm <= '"&t_yymm&"') and (pmg_id = '2') and (pmg_emp_no = '"&emp_no&"') and (pmg_company = '"&emp_company&"')"
Rs_give.Open Sql, Dbconn, 1
'Set Rs_give = DbConn.Execute(SQL)
do until Rs_give.eof
       pmg_yymm = Rs_give("pmg_yymm")
	   pay_year = mid(cstr(Rs_give("pmg_yymm")),1,4)
            pmg_give_tot = int(Rs_give("pmg_give_total"))	
	        sum_bunus_tot = sum_bunus_tot + pmg_give_tot

  		    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '2') and (de_emp_no = '"&emp_no&"') and (de_company = '"&emp_company&"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_nps_amt = int(Rs_dct("de_nps_amt"))	
					 de_nhis_amt = int(Rs_dct("de_nhis_amt"))	
					 de_epi_amt = int(Rs_dct("de_epi_amt"))	
					 de_longcare_amt = int(Rs_dct("de_longcare_amt"))	
					 de_income_tax = int(Rs_dct("de_income_tax"))	
					 de_wetax = int(Rs_dct("de_wetax"))	
                  else
                     de_nps_amt = 0
					 de_nhis_amt = 0
					 de_epi_amt = 0
					 de_longcare_amt = 0
					 de_income_tax = 0
					 de_wetax = 0
              end if
              Rs_dct.close()
			  sum_nps_amt = sum_nps_amt + de_nps_amt
	          sum_nhis_amt = sum_nhis_amt + de_nhis_amt
			  sum_epi_amt = sum_epi_amt + de_epi_amt
	          sum_longcare_amt = sum_longcare_amt + de_longcare_amt
			  sum_income_tax = sum_income_tax + de_income_tax
	          sum_wetax = sum_wetax + de_wetax
	Rs_give.MoveNext()
loop
Rs_give.close()

'종전근무지 자료 읽어서 현근무지 급여자료와 합해야 함.....
sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"'"
Rs_bef.Open Sql, Dbconn, 1
'Set Rs_bef = DbConn.Execute(SQL)
do until Rs_bef.eof
            b_pay = int(Rs_bef("b_pay"))	
			b_bonus = int(Rs_bef("b_bonus"))
			b_deem_bonus = int(Rs_bef("b_deem_bonus"))
			b_overtime_taxno = int(Rs_bef("b_overtime_taxno"))
			b_nps = int(Rs_bef("b_nps"))
			b_nhis = int(Rs_bef("b_nhis"))
			b_epi = int(Rs_bef("b_epi"))
			b_longcare = int(Rs_bef("b_longcare"))
			b_income_tax = int(Rs_bef("b_income_tax"))
			b_wetax = int(Rs_bef("b_wetax"))
			
            sum_give_tot = sum_give_tot + b_pay
			sum_bunus_tot = sum_bunus_tot + b_bonus + b_deem_bonus
	        sum_tax_no = sum_tax_no + b_overtime_taxno
			sum_nps_amt = sum_nps_amt + b_nps
	        sum_nhis_amt = sum_nhis_amt + b_nhis
		    sum_epi_amt = sum_epi_amt + b_epi
	        sum_longcare_amt = sum_longcare_amt + b_longcare
			sum_income_tax = sum_income_tax + b_income_tax
	        sum_wetax = sum_wetax + b_wetax
	Rs_bef.MoveNext()
loop
Rs_bef.close()			

title_line = " 소득자 정보 관리 "

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
			function getPageCode(){
				return "8 1";
			}
		</script>
	<script type="text/javascript">
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  

			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}		
			
			function chkfrm() {
//				if(document.frm.emp_ename.value =="") {
//					alert('영문성명을 입력하세요');
//					frm.emp_ename.focus();
//					return false;}
					
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.y_householder[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("세대주 구분을 체크하세요");
					return false;
				}	

				a=confirm('등록하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
			}
		</script>

</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
		  <div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax2_mg_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>성명(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td class="left"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (입사일:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th>소속(<%=emp_company%><input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true">)</th>
							  <td colspan="2" class="left"><%=emp_org_name%>
                                <input name="emp_org_name" type="hidden" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                - <%=emp_grade%>
                                <input name="emp_grade" type="hidden" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                - <%=emp_position%>
                                <input name="emp_position" type="hidden" value="<%=emp_position%>" style="width:70px" readonly="true">
                                (귀속년도:
                                <input name="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:40px; text-align:center" readonly="true">)
                              </td>
						    </tr>
                            <tr>
							  <th>세대주여부</th>
							  <td class="left">
							  <input type="radio" name="y_householder" value="Y" <% if y_householder = "Y" then %>checked<% end if %>>세대주
              		          <input name="y_householder" type="radio" value="N" <% if y_householder = "N" then %>checked<% end if %>>세대원
                              </td>
							  <th>거주구분/인정공제항목 변동여부</th>
							  <td colspan="2" class="left">
							  <input type="radio" name="y_live" value="Y" <% if y_live = "Y" then %>checked<% end if %>>거주자
              		          <input name="y_live" type="radio" value="N" <% if y_live = "N" then %>checked<% end if %>>비거주자
                              &nbsp;&nbsp; &nbsp;&nbsp;[인정공제항목 변동여부:&nbsp;&nbsp;
                              <input type="radio" name="y_change" value="N" <% if y_change = "N" then %>checked<% end if %>>전년과동일
              		          <input name="y_change" type="radio" value="Y" <% if y_change = "Y" then %>checked<% end if %>>변동&nbsp;]
                              </td>
						    </tr>
                            <tr>
							  <th>해외근로비과세</th>
                              <th>장애인사원</th>
                              <th>부녀자공제</th>
                              <th>한부모공제</th>
                              <th>생산직근로자</th>
						    </tr>
                            <tr>
							  <td class="center">
					          <input type="checkbox" name="y_foreign" value="Y" <% if y_foreign = "Y" then %>checked<% end if %> id="y_foreign">
                              </td>
                              <td class="center">
					          <input type="checkbox" name="y_disab" value="Y" <% if y_disab = "Y" then %>checked<% end if %> id="y_disab">
                              </td>
                              <td class="center">
					          <input type="checkbox" name="y_woman" value="Y" <% if y_woman = "Y" then %>checked<% end if %> id="y_woman">
                              </td>
                              <td class="center">
					          <input type="checkbox" name="y_single" value="Y" <% if y_single = "Y" then %>checked<% end if %> id="y_single">
                              </td>
                              <td class="center">
					          <input type="checkbox" name="y_blue" value="Y" <% if y_blue = "Y" then %>checked<% end if %> id="y_blue">
                              </td>
						    </tr>
						</tbody>
					</table>
				<h3 class="stit">※ 본인이 장애인인 경우는 인사팀에서 자동등록됨 / 본인이 기혼여성이거나 부양가족이 있는 세대주인 경우 부녀자 공제 대상으로 체크를 해야 함(남성 근로자는 해당없음)<br>
                ※ 배우자가 없고 기본공제대상인 직계비속 또는 입양자가 있는 경우 한부모공제 대상으로 체크(단 부녀자공제와 중복 적용안됨)</h3>
                <h3 class="stit">* 부양가족 및 당해년도 소득&nbsp;&nbsp;&nbsp;(소득금액은 이전근무지 소득과 합산해서 보여짐)</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                              <col width="4%" >
                              <col width="8%" >
                              <col width="10%" >
                              <col width="*" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">부양</th>
                                <th scope="col">관계</th>
                                <th scope="col">성명</th>
                                <th scope="col">주민등록번호</th>
                                <th scope="col">기본공제</th>
                                <th scope="col">경로우대</th>
                                <th scope="col">장애인</th>
                                <th scope="col">입양</th>
                                <th scope="col">자녀양육</th>
                                <th scope="col">위탁아동</th>
                                <th scope="col">수급자</th>
                              </tr>
                            </thead>
                            <tbody>
					<%
							sql = "select * from emp_family where family_empno = '"&emp_no&"' ORDER BY family_empno,family_seq ASC"
							Rs_fam.Open Sql, Dbconn, 1
							do until Rs_fam.eof
							    family_birthday = Rs_fam("family_birthday")
								family_support_yn = Rs_fam("family_support_yn")
								if family_birthday < "1944-12-31" then
								       family_old = "Y"
								   else
								       family_old = ""
								end if 
								if family_birthday > "2009-12-31" then
								       family_age6 = "Y"
								   else
								       family_age6 = ""
								end if 
								if family_support_yn = "Y" then
								       family_target = "Y"
								   else
								       family_target = ""
								end if   
						        family_disab = Rs_fam("family_disab")
						        family_merit = Rs_fam("family_merit")
						        family_serius = Rs_fam("family_serius")
						        family_witak = Rs_fam("family_witak")
						        family_holt = Rs_fam("family_holt")  
								family_pensioner = Rs_fam("family_pensioner")  
					%>
                                <tr>
                                  <td class="first"><input type="checkbox" name="support_check" value="Y" <% if family_support_yn = "Y" then %>checked<% end if %> id="support_check"></td>
                                  <td><%=Rs_fam("family_rel")%></td>
                                  <td><%=Rs_fam("family_name")%>&nbsp;</td>
                                  <td><%=Rs_fam("family_person1")%>-<%=Rs_fam("family_person2")%>&nbsp;</td>
                                  <td><input type="checkbox" name="family_target" value="Y" <% if family_target = "Y" then %>checked<% end if %> id="family_target"></td>
                                  <td><input type="checkbox" name="family_old" value="Y" <% if family_old = "Y" then %>checked<% end if %> id="family_old"></td>
                                  <td><input type="checkbox" name="family_disab" value="Y" <% if family_disab = "Y" then %>checked<% end if %> id="family_disab"></td>
                                  <td><input type="checkbox" name="family_holt" value="Y" <% if family_holt = "Y" then %>checked<% end if %> id="family_holt"></td>
                                  <td><input type="checkbox" name="family_age6" value="Y" <% if family_age6 = "Y" then %>checked<% end if %> id="family_age6"></td>
                                  <td><input type="checkbox" name="family_witak" value="Y" <% if family_witak = "Y" then %>checked<% end if %> id="family_witak"></td>
                                  <td><input type="checkbox" name="family_pensioner" value="Y" <% if family_pensioner = "Y" then %>checked<% end if %> id="family_pensioner"></td>
                              </tr>
							<%
								Rs_fam.movenext()
							loop
							Rs_fam.close()
							%>
                          </tbody>                        
                        </table>
                        </td>
                        <td width="2%"></td>
                        <td width="29%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                              <col width="15%" >
                              <col width="14%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">소득/공제</th>
                                <th scope="col">금액</th>
                              </tr>
                            </thead>
                            <tbody>
                                <tr>
                                  <td class="first">급여총액</td>
                                  <td class="right"><%=formatnumber(sum_give_tot,0)%>&nbsp;
                                  <input type="hidden" name="sum_give_tot" value="<%=formatnumber(sum_give_tot,0)%>" ID="sum_give_tot"></td>
                                </tr> 
                                <tr> 
                                  <td class="first">상여총액</td>
                                  <td class="right"><%=formatnumber(sum_bunus_tot,0)%>&nbsp;
                                  <input type="hidden" name="sum_bunus_tot" value="<%=formatnumber(sum_bunus_tot,0)%>" ID="sum_bunus_tot"></td>
                                </tr>
                                <tr> 
                                  <td class="first">기타소득</td>
                                  <td class="right"><%=formatnumber(sum_other_tot,0)%>&nbsp;
                                  <input type="hidden" name="sum_other_tot" value="<%=formatnumber(sum_other_tot,0)%>" ID="sum_other_tot"></td>
                                </tr> 
                                <tr> 
                                  <td class="first">비과세</td>
                                  <td class="right"><%=formatnumber(sum_tax_no,0)%>&nbsp;
                                  <input type="hidden" name="sum_tax_no" value="<%=formatnumber(sum_tax_no,0)%>" ID="sum_tax_no"></td>
                                </tr> 
                                <tr> 
                                  <td class="first">기납부소득세</td>
                                  <td class="right"><%=formatnumber(sum_income_tax,0)%>&nbsp;
                                  <input type="hidden" name="sum_income_tax" value="<%=formatnumber(sum_income_tax,0)%>" ID="sum_income_tax"></td>
                                </tr> 
                                <tr> 
                                  <td class="first">기납부주민세</td>
                                  <td class="right"><%=formatnumber(sum_wetax,0)%>&nbsp;
                                  <input type="hidden" name="sum_wetax" value="<%=formatnumber(sum_wetax,0)%>" ID="sum_wetax"></td>
                                </tr>
                                <tr> 
                                  <td class="first">국민연금</td>
                                  <td class="right"><%=formatnumber(sum_nps_amt,0)%>&nbsp;
                                  <input type="hidden" name="sum_nps_amt" value="<%=formatnumber(sum_nps_amt,0)%>" ID="sum_nps_amt"></td>
                                </tr>
                                <tr> 
                                  <td class="first">건강보험</td>
                                  <td class="right"><%=formatnumber(sum_nhis_amt,0)%>&nbsp;
                                  <input type="hidden" name="sum_nhis_amt" value="<%=formatnumber(sum_nhis_amt,0)%>" ID="sum_nhis_amt"></td>
                                </tr>
                                <tr> 
                                  <td class="first">요양보험</td>
                                  <td class="right"><%=formatnumber(sum_longcare_amt,0)%>&nbsp;
                                  <input type="hidden" name="sum_longcare_amt" value="<%=formatnumber(sum_longcare_amt,0)%>" ID="sum_longcare_amt"></td>
                                </tr>
                                <tr> 
                                  <td class="first">고용보험</td>
                                  <td class="right"><%=formatnumber(sum_epi_amt,0)%>&nbsp;
                                  <input type="hidden" name="sum_epi_amt" value="<%=formatnumber(sum_epi_amt,0)%>" ID="sum_epi_amt"></td>
                                </tr>  
                            </tbody>
                        </table>
                        </td>
                      </tr>
              </table>
				<br>
                <div align=center>
                <% if y_final <> "Y" then  %>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                <% end if  %>
                    <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();"></span>
                </div>
                    <input type="hidden" name="company_name" value="<%=company_name%>" ID="company_name">
			        <input type="hidden" name="trade_no" value="<%=trade_no%>" ID="trade_no">
                    <input type="hidden" name="emp_national" value="<%=emp_national%>" ID="emp_national">
              </form>                    
		  </div>
		</div>				
	</div>        				
	</body>
</html>

