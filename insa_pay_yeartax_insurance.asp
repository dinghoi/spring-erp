<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_insurance.asp"

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

y_nhis_amt = 0
y_longcare_amt = 0
y_epi_amt = 0
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
Set rs_year = DbConn.Execute(SQL)
if not rs_year.eof then
       y_nhis_amt = rs_year("y_nhis_amt")
	   y_longcare_amt = rs_year("y_longcare_amt")
	   y_epi_amt = rs_year("y_epi_amt")
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
Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
       b_nhis = b_nhis + rs_bef("b_nhis")
	   b_longcare = b_longcare + rs_bef("b_longcare")
	   b_epi = b_epi + rs_bef("b_epi")
	rs_bef.MoveNext()
loop
rs_bef.close()
b_nhis = b_nhis + b_longcare
b_nhis_tax = b_nhis
b_epi_tax = b_epi

ilban_insu = 0
disab_insu = 0
Sql = "select * from pay_yeartax_insurance where i_year = '"&inc_yyyy&"' and i_emp_no = '"&emp_no&"' ORDER BY i_emp_no,i_seq ASC"
rs_ins.Open Sql, Dbconn, 1
Set rs_ins = DbConn.Execute(SQL)
do until rs_ins.eof
       if rs_ins("i_disab_chk") = "Y" then
	          disab_insu = disab_insu + rs_ins("i_nts_amt") + rs_ins("i_other_amt")
		  else	  
			  ilban_insu = ilban_insu + rs_ins("i_nts_amt") + rs_ins("i_other_amt")
	   end if
	rs_ins.MoveNext()
loop
rs_ins.close()

if ilban_insu > 1000000 then 
       ilban_insu_tax = 1000000
   else
       ilban_insu_tax = ilban_insu
end if

if disab_insu > 1000000 then 
       disab_insu_tax = 1000000
   else
       disab_insu_tax = disab_insu
end if

tot_amt = y_nhis_amt + b_nhis + y_epi_amt + b_epi + ilban_insu + disab_insu
tot_tax = y_nhis_tax + b_nhis_tax + y_epi_tax + b_epi_tax + ilban_insu_tax + disab_insu_tax

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
Set rs_fami = DbConn.Execute(SQL)
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

sql = "select * from pay_yeartax_insurance where i_year = '"&inc_yyyy&"' and i_emp_no = '"&emp_no&"' ORDER BY i_emp_no,i_person_no,i_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 특별공제(보험료) "
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
				<form action="insa_pay_yeartax_insurance.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="14%" >
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
							  <th rowspan="7">보험료</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3;">국민건강보험<br>(노인장기요양보험 포함)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(b_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(b_nhis_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right" style=" border-bottom:1px solid #e3e3e3;"><%=formatnumber(y_nhis_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nhis_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">고용보험</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(b_epi,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(b_epi_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(y_epi_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_epi_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">일반보장성 보험</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(ilban_insu,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;100만원</th>
                              <td class="right"><%=formatnumber(ilban_insu_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장애인전용보장성보험</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(disab_insu,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">100만원</th>
                              <td class="right"><%=formatnumber(disab_insu_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;">보험료 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_tax,0)%>&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">※ 근로소득자가 지출한 경비중 사회보장성 경비인 건강보험료, 고용보험료,노인장기요양보험료 전액과 보장성보험료 및 장애인전용보장성보험료 각 연 100만원 한도내의 금액을 근로소득에서 공제<br>
                ※ 보험계약자 및 피보험자(보험대상자) 모두 인적공제에 기본공제대상이어야 함<br>
                ※ 장애인전용 보장성 보험료는 별도로 인사팀으로 직접 연락 하셔야 합니다.</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="16%" >
                              <col width="16%" >
                              <col width="16%" >
                              <col width="12%" >
                              <col width="16%" >
                              <col width="16%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">선택</th>
                                <th scope="col">관계(유형)</th>
                                <th scope="col">대상자이름</th>
                                <th scope="col">주민등록번호</th>
                                <th scope="col">장애인전용<br>보장성보험여부</th>
                                <th scope="col">국세청금액</th>
                                <th scope="col">그밖의금액</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             i_disab_chk = rs("i_disab_chk")
	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("i_rel")%>&nbsp;</td>
                                <td><%=rs("i_name")%>&nbsp;</td>
                                <td><%=mid(cstr(rs("i_person_no")),1,6)%>-<%=mid(cstr(rs("i_person_no")),7,7)%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="i_disab_chk" value="Y" <% if i_disab_chk = "Y" then %>checked<% end if %> id="i_disab_chk"></td>
                                <td class="right"><%=formatnumber(rs("i_nts_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("i_other_amt"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                      
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_insurance_add.asp?i_year=<%=rs("i_year")%>&i_emp_no=<%=rs("i_emp_no")%>&i_seq=<%=rs("i_seq")%>&i_person_no=<%=rs("i_person_no")%>&i_emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_pay_yeartax_insurance_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
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
              <% if y_final <> "Y" then  %>                        
					<a href="#" onClick="pop_Window('insa_pay_yeartax_insurance_add.asp?i_year=<%=inc_yyyy%>&i_emp_no=<%=emp_no%>&i_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_insurance_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">보험료추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_insurance.asp" class="btnType04">보험료등록</a>
			  <%   end if  %>                        
                    <a href="insa_pay_yeartax_medical.asp" class="btnType04">의료비등록</a>
                    <a href="insa_pay_yeartax_edu.asp" class="btnType04">교육비등록</a>
                    <a href="insa_pay_yeartax_house.asp" class="btnType04">주택자금등록</a>
                    <a href="insa_pay_yeartax_donation.asp" class="btnType04">기부금등록</a>
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

