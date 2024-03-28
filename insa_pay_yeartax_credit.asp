<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_credit.asp"

y_final=Request("y_final")
c_id=Request("c_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	inc_yyyy = request.form("inc_yyyy")
  else
	inc_yyyy = request("inc_yyyy")
end if

if view_condi = "" then
	'inc_yyyy = mid(cstr(now()),1,4)
	inc_yyyy = cint(mid(now(),1,4)) - 1
	ck_sw = "n"
end if

for i = 1 to 10
    family_tab(i,1) = ""
	family_tab(i,2) = ""
	family_tab(i,3) = ""
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
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
emp_person = cstr(rs_emp("emp_person1")) + cstr(rs_emp("emp_person2"))	
rs_emp.close()	

tot_pay = 0
y_nhis_amt = 0
y_longcare_amt = 0
y_epi_amt = 0
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
'Set rs_year = DbConn.Execute(SQL)
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
y_nhis_amt = y_nhis_amt + y_longcare_amt
y_nhis_tax = y_nhis_amt
y_epi_tax = y_epi_amt

b_nhis = 0
b_longcare = 0
b_epi = 0
Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
'Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
       b_nhis = b_nhis + rs_bef("b_nhis")
	   b_longcare = b_longcare + rs_bef("b_longcare")
	   b_epi = b_epi + rs_bef("b_epi")
	   tot_pay = tot_pay + rs_bef("b_pay") + rs_bef("b_bonus") + rs_bef("b_deem_bonus")
	rs_bef.MoveNext()
loop
rs_bef.close()
b_nhis = b_nhis + b_longcare
b_nhis_tax = b_nhis
b_epi_tax = b_epi

tot_3per = int(tot_pay * (3 / 100))

market_tot = 0
transit_tot = 0
credit_tot = 0
jik_tot = 0
hyun_tot = 0
Sql = "select * from pay_yeartax_credit where c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' ORDER BY c_emp_no,c_person_no,c_id,c_seq ASC"
rs_cred.Open Sql, Dbconn, 1
'Set rs_cred = DbConn.Execute(SQL)
do until rs_cred.eof
       if rs_cred("c_market") = "Y" then
	           market_tot = market_tot + rs_cred("c_nts_amt") + rs_cred("c_other_amt")
		  elseif rs_cred("c_transit") = "Y" then
			           transit_tot = transit_tot + rs_cred("c_nts_amt") + rs_cred("c_other_amt")
				  elseif rs_cred("c_id") = "신용카드" then
				              credit_tot = credit_tot + rs_cred("c_nts_amt") + rs_cred("c_other_amt")
					     elseif rs_cred("c_id") = "직불카드" then
				                     jik_tot = jik_tot + rs_cred("c_nts_amt") + rs_cred("c_other_amt") 
							    elseif rs_cred("c_id") = "직불카드" then
				                          hyun_tot = hyun_tot + rs_cred("c_nts_amt") + rs_cred("c_other_amt") 
		end if
	rs_cred.MoveNext()
loop
rs_cred.close()

tot_amt = market_tot + transit_tot + credit_tot + jik_tot + hyun_tot 
market_tax = market_tot
transit_tax = transit_tot
credit_tax = credit_tot
jik_tax = jik_tot
hyun_tax = hyun_tot
tot_tax = market_tax + transit_tax + credit_tax + jik_tax + hyun_tax

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
'Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "본인" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
		  i = i + 1
		  family_tab(i,1) = rs_fami("f_rel")
	      family_tab(i,2) = rs_fami("f_family_name")
	      family_tab(i,3) = rs_fami("f_person_no")
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

sql = "select * from pay_yeartax_credit where c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' ORDER BY c_emp_no,c_person_no,c_id,c_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 신용카드 사용액(" + c_id + ")"
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
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_credit.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="16%" >
							<col width="14%" >
							<col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">성명(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td colspan="2" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (입사일:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th style=" border-bottom:1px solid #e3e3e3;">소속(<%=emp_company%><input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true">)</th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_org_name%>
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
							  <th style=" border-bottom:1px solid #e3e3e3;">구분</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">지출명세</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">지출구분</th>
                              <th>금액</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">한도액</th>
                              <th>공제액</th>
						    </tr>
                            <tr>
							  <th rowspan="6">신용카드</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">신용카드(전통시장·대중교통사용분 제외)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">사용금액</th>
                              <td class="right"><%=formatnumber(credit_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" style="background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">직불·선불카드(전통시장·대중교통사용분 제외)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">사용금액</th>
                              <td class="right"><%=formatnumber(jik_tax,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" style="background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">현금영수증(전통시장·대중교통사용분 제외)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">사용금액</th>
                              <td class="right"><%=formatnumber(hyun_tax,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" style="background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">전통시장사용분</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">사용금액</th>
                              <td class="right"><%=formatnumber(market_tax,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" style="background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">대중교통이용분</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">사용금액</th>
                              <td class="right"><%=formatnumber(transit_tax,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" style="background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;">신용카드 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_tax,0)%>&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">※ 근로자본인과 연간 소득금액이 100만원 이하인 배우자와 직계존비속(나이제한 없음)의 신용카드 등 사용금액만 공제대상이 됨.<br>
                ※ 형제,자매의 신용카드등의 사용금액은 공제대상이 아닙니다. 절대 입력하지 마세요.<br>
                ※ 신용카드,직불카드,현금영수증 사용분중 전통시장 및 대중교통에 해당하는 금액은 전통시장 또는 대중교통에 체크하고 입력.</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="*" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">선택</th>
                                <th scope="col">사용구분</th>
                                <th scope="col">관계(유형)</th>
                                <th scope="col">대상자이름</th>
                                <th scope="col">주민등록번호</th>
                                <th scope="col">전통시장</th>
                                <th scope="col">대중교통</th>
                                <th scope="col">국세청 금액</th>
                                <th scope="col">기타 금액</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             c_market = rs("c_market")
							 c_transit = rs("c_transit")
	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=c_id%>&nbsp;</td>
                                <td><%=rs("c_rel")%>&nbsp;</td>
                                <td><%=rs("cc_name")%>&nbsp;</td>
                                <td><%=mid(cstr(rs("c_person_no")),1,6)%>-<%=mid(cstr(rs("c_person_no")),7,7)%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="c_market" value="Y" <% if c_market = "Y" then %>checked<% end if %> id="c_market"></td>
                                <td>
                                <input type="checkbox" name="c_transit" value="Y" <% if c_transit = "Y" then %>checked<% end if %> id="c_transit"></td>
                                <td class="right"><%=formatnumber(rs("c_nts_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("c_other_amt"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                   
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_credit_add.asp?c_year=<%=rs("c_year")%>&c_emp_no=<%=rs("c_emp_no")%>&c_seq=<%=rs("c_seq")%>&c_person_no=<%=rs("c_person_no")%>&c_emp_name=<%=emp_name%>&c_id=<%=c_id%>&u_type=<%="U"%>','insa_pay_yeartax_credit_add_pop','scrollbars=yes,width=850,height=370')">수정</a></td>
                        <%    else  %>
                                <td>&nbsp;</td>
                        <% end if  %>									
                            </tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
              <% if c_id = "신용카드" then  %>
					<div class="btnRight">
              <% if y_final <> "Y" then  %>                        
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_credit_add.asp?c_year=<%=inc_yyyy%>&c_emp_no=<%=emp_no%>&c_emp_name=<%=emp_name%>&c_id=<%=c_id%>&u_type=<%=""%>','insa_pay_yeartax_edu_add_pop','scrollbars=yes,width=850,height=370')" class="btnType04">신용카드추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="신용카드"%>" class="btnType04">신용카드</a>
			  <%   end if  %>                     
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="직불카드"%>" class="btnType04">직불카드</a>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="현금영수증"%>" class="btnType04">현금영수증</a>
					</div> 
			  <% end if %>		
              <% if c_id = "직불카드" then  %>  
                    <div class="btnRight">
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="신용카드"%>" class="btnType04">신용카드</a>
              <% if y_final <> "Y" then  %>                       
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_credit_add.asp?c_year=<%=inc_yyyy%>&c_emp_no=<%=emp_no%>&c_emp_name=<%=emp_name%>&c_id=<%=c_id%>&u_type=<%=""%>','insa_pay_yeartax_edu_add_pop','scrollbars=yes,width=850,height=370')" class="btnType04">직불카드추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="직불카드"%>" class="btnType04">직불카드</a>
			  <%   end if  %>                      
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="현금영수증"%>" class="btnType04">현금영수증</a>
					</div> 
              <% end if %>		
              <% if c_id = "현금영수증" then  %> 
                    <div class="btnRight">
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="신용카드"%>" class="btnType04">신용카드</a>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="직불카드"%>" class="btnType04">직불카드</a>
              <% if y_final <> "Y" then  %>                      
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_credit_add.asp?c_year=<%=inc_yyyy%>&c_emp_no=<%=emp_no%>&c_emp_name=<%=emp_name%>&c_id=<%=c_id%>&u_type=<%=""%>','insa_pay_yeartax_edu_add_pop','scrollbars=yes,width=850,height=370')" class="btnType04">현금영수증추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="현금영수증"%>" class="btnType04">현금영수증</a>
			  <%   end if  %>                      
					</div>  
              <% end if %>	
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">                
			</form>
		</div>				
	</div>        				
	</body>
</html>

