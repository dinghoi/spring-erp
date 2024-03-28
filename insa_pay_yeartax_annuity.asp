<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_annuity.asp"

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

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ann = Server.CreateObject("ADODB.Recordset")
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
rs_emp.close()	

y_nps_other = 0
y_nps_amt = 0
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
Set rs_year = DbConn.Execute(SQL)
if not rs_year.eof then
       y_nps_amt = rs_year("y_nps_amt")
   else
       y_nps_amt = 0
end if
y_nps_tax = y_nps_amt

b_nps = 0
Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
       b_nps = b_nps + rs_bef("b_nps")
	rs_bef.MoveNext()
loop
rs_bef.close()
b_nps_tax = b_nps

a_amt_other = 0
a_amt_tot = 0
Sql = "select * from pay_yeartax_annuity where a_year = '"&inc_yyyy&"' and a_emp_no = '"&emp_no&"' ORDER BY a_emp_no,a_seq ASC"
rs_ann.Open Sql, Dbconn, 1
Set rs_ann = DbConn.Execute(SQL)
do until rs_ann.eof
       a_amt_tot = a_amt_tot + rs_ann("a_amt")
	rs_ann.MoveNext()
loop
rs_ann.close()

a_amt_tax = a_amt_tot

a_amt_other_tax = a_amt_other
y_nps_other_tax = y_nps_other

tot_amt = y_nps_amt + b_nps + y_nps_other + a_amt_other + a_amt_tot
tot_tax = y_nps_tax + b_nps_tax + y_nps_other_tax + a_amt_other_tax + a_amt_tax


sql = "select * from pay_yeartax_annuity where a_year = '"&inc_yyyy&"' and a_emp_no = '"&emp_no&"' ORDER BY a_emp_no,a_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 연금보험공제 "
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
				<form action="insa_pay_yeartax_annuity.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
							  <th rowspan="7">연금보험료<br>(국민연금,공무원연금,군인연금,교직원연금,연금계좌등)</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3;">국민연금보험료</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(b_nps,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(b_nps_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right" style=" border-bottom:1px solid #e3e3e3;"><%=formatnumber(y_nps_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nps_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">국민연금보험료 외의 공적연금보험료</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(y_nps_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nps_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(y_nps_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nps_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">연금계좌</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(a_amt_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">작성방법참조</th>
                              <td class="right"><%=formatnumber(a_amt_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(a_amt_tot,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">작성방법참조</th>
                              <td class="right"><%=formatnumber(a_amt_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;">연금보험료 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_tax,0)%>&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">※ 금융기관/계좌번호 or 증권번호를 정확하게 입력<br>
                ※ 퇴직연금은 확정기여형퇴직연금(DC)형과 개인형퇴직연금계좌(IRP)에 본인이 추가로 불입한 금액에 대하여 입력</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="20%" >
                              <col width="16%" >
                              <col width="20%" >
                              <col width="20%" >
                              <col width="16%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">선택</th>
                                <th scope="col">유형</th>
                                <th scope="col">금융기관</th>
                                <th scope="col">금융사명</th>
                                <th scope="col">계좌/증권번호</th>
                                <th scope="col">금액</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("a_type")%>&nbsp;</td>
                                <td><%=rs("a_bank_code")%>&nbsp;</td>
                                <td><%=rs("a_bank_name")%>&nbsp;</td>
                                <td><%=rs("a_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("a_amt"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_annuity_add.asp?a_year=<%=rs("a_year")%>&a_emp_no=<%=rs("a_emp_no")%>&a_seq=<%=rs("a_seq")%>&a_emp_name=<%=rs("a_emp_name")%>&u_type=<%="U"%>','insa_pay_yeartax_annuity_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
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
              <% if y_final <> "Y" then  %>                    
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_yeartax_annuity_add.asp?a_year=<%=inc_yyyy%>&a_emp_no=<%=emp_no%>&a_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_annuity_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">연금보험료 세부항목입력</a>
					</div>                  
              <%   else  %>
                       <br><br>
			  <%   end if  %>                       
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">                
			</form>
		</div>				
	</div>        				
	</body>
</html>

