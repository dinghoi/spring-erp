<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_edu.asp"

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
Set rs_edu = Server.CreateObject("ADODB.Recordset")
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
Set rs_year = DbConn.Execute(SQL)
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
Set rs_bef = DbConn.Execute(SQL)
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

bon_tot = 0
adong_tot = 0
school_tot = 0
univ_tot = 0
disab_tot = 0

adong_cnt = 0
school_cnt = 0
univ_cnt = 0
disab_cnt = 0
Sql = "select * from pay_yeartax_edu where e_year = '"&inc_yyyy&"' and e_emp_no = '"&emp_no&"' ORDER BY e_emp_no,e_person_no,e_seq ASC"
rs_edu.Open Sql, Dbconn, 1
Set rs_edu = DbConn.Execute(SQL)
do until rs_edu.eof
       if rs_edu("e_edu_level") = "장애인" then
	         disab_tot = disab_tot + rs_edu("e_nts_amt") + rs_edu("e_other_amt")
			 disab_cnt = disab_cnt + 1
		  elseif rs_edu("e_edu_level") = "소득자본인" then
		             bon_tot = bon_tot + rs_edu("e_nts_amt") + rs_edu("e_other_amt")
				 elseif rs_edu("e_edu_level") = "취학전아동" then
	                       adong_tot = adong_tot + rs_edu("e_nts_amt") + rs_edu("e_other_amt")
						   adong_cnt = adong_cnt + 1
						elseif rs_edu("e_edu_level") = "초/중/고등학교" then
	                              school_tot = school_tot + rs_edu("e_nts_amt") + rs_edu("e_other_amt")
								  school_cnt = school_cnt + 1
						       elseif rs_edu("e_edu_level") = "대학생(대학원불포함)" then
	                                     univ_tot = univ_tot + rs_edu("e_nts_amt") + rs_edu("e_other_amt") 
										 univ_cnt = univ_cnt + 1
		end if
	rs_edu.MoveNext()
loop
rs_edu.close()

tot_amt = bon_tot + adong_tot + school_tot + univ_tot + disab_tot 
bon_tax = bon_tot
adong_tax = adong_tot
school_tax = school_tot
univ_tax = univ_tot
disab_tax = disab_tot
tot_tax = bon_tax + adong_tax + school_tax + univ_tax + disab_tot

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

sql = "select * from pay_yeartax_edu where e_year = '"&inc_yyyy&"' and e_emp_no = '"&emp_no&"' ORDER BY e_emp_no,e_person_no,e_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 특별공제(교육비) "
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
				<form action="insa_pay_yeartax_edu.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
							  <th rowspan="6">의료비</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">소득자 본인</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공납금(대학원포함)</th>
                              <td class="right"><%=formatnumber(bon_tot,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(bon_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">취학전 아동(<%=adong_cnt%>명)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">유치원·학원비등</th>
                              <td class="right"><%=formatnumber(adong_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">1명당 300만원</th>
                              <td class="right"><%=formatnumber(adong_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">초·중·고등학교(<%=school_cnt%>명)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">공납금</th>
                              <td class="right"><%=formatnumber(school_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">1명당 300만원</th>
                              <td class="right"><%=formatnumber(school_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">대학생(대학원불포함)(<%=univ_cnt%>명)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">공납금</th>
                              <td class="right"><%=formatnumber(univ_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">1명당 900만원</th>
                              <td class="right"><%=formatnumber(univ_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장애인(<%=disab_cnt%>명)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">특수교육비</th>
                              <td class="right"><%=formatnumber(disab_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(disab_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;">교육비 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_tax,0)%>&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">※ 대학원·시간제과정·직업능력개발 교육비는 본인에 한하여 공제가능.<br>
                ※ 국내에서 송금한 국외 교육비의 경우,해외 송금일의 대고객 외국환매도율을 적용하여 원화로 환산한 교육비금액을 입력.<br>
                ※ 교육비 공제의 교육비는 근로자의 교육비 또는 기본공제대상자(연령요건 불문하나 직계존속은 제외함)을 위하여 지출한 금액입니다.</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="*" >
                              <col width="8%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">선택</th>
                                <th scope="col">관계(유형)</th>
                                <th scope="col">대상자이름</th>
                                <th scope="col">주민등록번호</th>
                                <th scope="col">장애<br>여부</th>
                                <th scope="col">교육수준</th>
                                <th scope="col">교복구입비<br>여부</th>
                                <th scope="col">국세청 금액</th>
                                <th scope="col">기타 금액</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             e_disab = rs("e_disab")
							 e_uniform = rs("e_uniform")
	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("e_rel")%>&nbsp;</td>
                                <td><%=rs("e_name")%>&nbsp;</td>
                                <td><%=mid(cstr(rs("e_person_no")),1,6)%>-<%=mid(cstr(rs("e_person_no")),7,7)%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="e_disab" value="Y" <% if e_disab = "Y" then %>checked<% end if %> id="e_disab"></td>
                                <td><%=rs("e_edu_level")%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="e_uniform" value="Y" <% if e_uniform = "Y" then %>checked<% end if %> id="e_uniform"></td>
                                <td class="right"><%=formatnumber(rs("e_nts_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("e_other_amt"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                     
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_edu_add.asp?e_year=<%=rs("e_year")%>&e_emp_no=<%=rs("e_emp_no")%>&e_seq=<%=rs("e_seq")%>&e_person_no=<%=rs("e_person_no")%>&e_emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_pay_yeartax_edu_add_pop','scrollbars=yes,width=850,height=370')">수정</a></td>
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
              <% if y_final <> "Y" then  %>                        
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_edu_add.asp?e_year=<%=inc_yyyy%>&e_emp_no=<%=emp_no%>&e_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_edu_add_pop','scrollbars=yes,width=850,height=370')" class="btnType04">교육비추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_edu.asp" class="btnType04">교육비등록</a>
			  <%   end if  %>                               
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

