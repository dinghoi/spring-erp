<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_donation.asp"

y_final=Request("y_final")

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
Set rs_dona = Server.CreateObject("ADODB.Recordset")
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

p_tot = 0 '정치
l_tot = 0 '법정
e_tot = 0 '조합
ro_tot = 0 '종교외
r_tot = 0 '종교
Sql = "select * from pay_yeartax_donation where d_year = '"&inc_yyyy&"' and d_emp_no = '"&emp_no&"' ORDER BY d_emp_no,d_person_no,d_seq ASC"
rs_dona.Open Sql, Dbconn, 1
'Set rs_dona = DbConn.Execute(SQL)
do until rs_dona.eof
       if rs_dona("d_data_gubun") = "정치자금기부금" then
	         p_tot = p_tot + rs_dona("d_amt") 
		  elseif rs_dona("d_data_gubun") = "법정기부금" then
		             l_tot = l_tot + rs_dona("d_amt") 
				 elseif rs_dona("d_data_gubun") = "우리사주조합기부금" then
	                       e_tot = e_tot + rs_dona("d_amt") 
						elseif rs_dona("d_data_gubun") = "종교단체외지정기부금" then
	                              ro_tot = ro_tot + rs_dona("d_amt") 
						       elseif rs_dona("d_data_gubun") = "종교단체지정기부금" then
	                                     r_tot = r_tot + rs_dona("d_amt") 
		end if
	rs_dona.MoveNext()
loop
rs_dona.close()

tot_amt = p_tot + l_tot + e_tot + ro_tot + r_tot
p_tax = p_tot
l_tax = l_tot
e_tax = e_tot
ro_tax = ro_tot
r_tax = r_tot
tot_tax = p_tax + l_tax + e_tax + ro_tax + r_tax 

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

sql = "select * from pay_yeartax_donation where d_year = '"&inc_yyyy&"' and d_emp_no = '"&emp_no&"' ORDER BY d_emp_no,d_person_no,d_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 특별공제(기부금) "
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
				<form action="insa_pay_yeartax_donation.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
							  <th rowspan="6">기부금</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">정치자금기부금(세액공제분 제외)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">기부금액</th>
                              <td class="right"><%=formatnumber(p_tot,0)%>&nbsp;</td>
                              <th rowspan="5" style=" border-bottom:1px solid #e3e3e3;">작성방법 참조</th>
                              <td class="right"><%=formatnumber(p_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">법정 기부금</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">기부금액</th>
                              <td class="right"><%=formatnumber(l_tot,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(l_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">우리사주조합 기부금</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">기부금</th>
                              <td class="right"><%=formatnumber(e_tot,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(e_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">종교단체 외 지정기부금</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">기부금</th>
                              <td class="right"><%=formatnumber(ro_tot,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(ro_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">종교단체 지정기부금</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">기부금</th>
                              <td class="right"><%=formatnumber(r_tot,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(r_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;">기부금공제 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_tax,0)%>&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">※ 정치자금 및 우리사주조합 기부금은 기부자가 근로자 본인일 때만 공제가능.<br>
                ※ 기부금유형과과 기부처가 같은경우는 기부금액을 합산하여 1건으로 입력.<br>
                ※ 기본공제 대상자 배우자 및 부양가족(직계존비속 포함)이 지급한 기부금도 공제대상임.</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="8%" >
                              <col width="10%" >
                              <col width="12%" >
                              <col width="12%" >
                              <col width="10%" >
                              <col width="*" >
                              <col width="8%" >
                              <col width="10%" >
                              <col width="5%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">선택</th>
                                <th scope="col">관계(유형)</th>
                                <th scope="col">대상자이름</th>
                                <th scope="col">주민등록번호</th>
                                <th scope="col">기부코드(유형)</th>
                                <th scope="col">사업자(주민)번호</th>
                                <th scope="col">기부처명</th>
                                <th scope="col">건수</th>
                                <th scope="col">금액</th>
                                <th scope="col">국세청<br>자료여부</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             d_nts_chk = rs("d_nts_chk")

	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("d_rel")%>&nbsp;</td>
                                <td><%=rs("d_name")%>&nbsp;</td>
                                <td><%=mid(cstr(rs("d_person_no")),1,6)%>-<%=mid(cstr(rs("d_person_no")),7,7)%>&nbsp;</td>
                                <td><%=rs("d_data_gubun")%>&nbsp;</td>
                                <td><%=rs("d_trade_no")%>&nbsp;</td>
                                <td><%=rs("d_trade_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("d_cnt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("d_amt"),0)%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="d_nts_chk" value="Y" <% if d_nts_chk = "Y" then %>checked<% end if %> id="d_nts_chk"></td>
                        <% if y_final <> "Y" then  %>                                 
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_donation_add.asp?d_year=<%=rs("d_year")%>&d_emp_no=<%=rs("d_emp_no")%>&d_seq=<%=rs("d_seq")%>&d_person_no=<%=rs("d_person_no")%>&d_emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_pay_yeartax_donation_add_pop','scrollbars=yes,width=850,height=350')">수정</a></td>
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
					<div class="btnRight">
                    <a href="insa_pay_yeartax_insurance.asp" class="btnType04">보험료등록</a>
                    <a href="insa_pay_yeartax_medical.asp" class="btnType04">의료비등록</a>
                    <a href="insa_pay_yeartax_edu.asp" class="btnType04">교육비등록</a>
                    <a href="insa_pay_yeartax_house.asp" class="btnType04">주택자금등록</a>
              <% if y_final <> "Y" then  %>                        
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_donation_add.asp?d_year=<%=inc_yyyy%>&d_emp_no=<%=emp_no%>&d_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_donation_add_pop','scrollbars=yes,width=850,height=350')" class="btnType04">기부금추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_donation.asp" class="btnType04">기부금등록</a>
			  <%   end if  %>   					
                    </div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">                
			</form>
		</div>				
	</div>        				
	</body>
</html>

