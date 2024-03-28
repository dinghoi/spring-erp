<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
b_year = request("b_year")
b_emp_no = request("b_emp_no")
b_emp_name = request("b_emp_name")
b_seq = request("b_seq")

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 이전근무지 정보 등록 "
if u_type = "U" then

	Sql="select * from pay_yeartax_before where b_year = '"&b_year&"' and b_emp_no = '"&b_emp_no&"' and b_seq = '"&b_seq&"'"
	Set rs=DbConn.Execute(Sql)

	b_emp_name = rs("b_emp_name")
    b_company_no = rs("b_company_no")
    b_company = rs("b_company")
    b_from_date = rs("b_from_date")
    b_to_date = rs("b_to_date")
    b_pay = rs("b_pay")
    b_bonus = rs("b_bonus")
    b_deem_bonus = rs("b_deem_bonus")
	b_foreign_taxno = rs("b_foreign_taxno")
    b_overtime_taxno = rs("b_overtime_taxno")
    b_age6 = rs("b_age6")
	b_nps = rs("b_nps")
	b_nhis = rs("b_nhis")
    b_epi = rs("b_epi")
	b_longcare = rs("b_longcare")
    b_income_tax = rs("b_income_tax")
    b_wetax = rs("b_wetax")
    b_stock_profit = rs("b_stock_profit")

	rs.close()

	title_line = " 이전근무지 정보 변경 "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=b_from_date%>" );
			});	
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=b_to_date%>" );
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
				if(document.frm.b_company_no.value =="") {
					alert('사업자등록번호를 입력하세요');
					frm.b_company_no.focus();
					return false;}
				if(document.frm.b_company.value =="") {
					alert('근무처명을 입력하세요');
					frm.b_company.focus();
					return false;}
				if(document.frm.b_from_date.value =="") {
					alert('근무시작일을 입력하세요');
					frm.b_from_date.focus();
					return false;}
				if(document.frm.b_to_date =="") {
					alert('근무종료일을 선택하세요');
					frm.b_to_date.focus();
					return false;}
				if(document.frm.b_pay.value =="") {
					alert('급여를 입력하세요');
					frm.b_pay.focus();
					return false;}
			
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			} 
			
			function num_chk(txtObj){
				bb_pay = parseInt(document.frm.b_pay.value.replace(/,/g,""));	
				bb_bonus = parseInt(document.frm.b_bonus.value.replace(/,/g,""));	
				deem_bonus = parseInt(document.frm.b_deem_bonus.value.replace(/,/g,""));	
				overtime_taxno = parseInt(document.frm.b_overtime_taxno.value.replace(/,/g,""));	
				
				bb_nps = parseInt(document.frm.b_nps.value.replace(/,/g,""));	
				bb_nhis = parseInt(document.frm.b_nhis.value.replace(/,/g,""));	
				bb_epi = parseInt(document.frm.b_epi.value.replace(/,/g,""));	
				bb_longcare = parseInt(document.frm.b_longcare.value.replace(/,/g,""));	
				
				bb_income_tax = parseInt(document.frm.b_income_tax.value.replace(/,/g,""));	
				bb_wetax = parseInt(document.frm.b_wetax.value.replace(/,/g,""));	
		
				bb_pay = String(bb_pay);
				num_len = bb_pay.length;
				sil_len = num_len;
				bb_pay = String(bb_pay);
				if (bb_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_pay = bb_pay.substr(0,num_len -3) + "," + bb_pay.substr(num_len -3,3);
				if (sil_len > 6) bb_pay = bb_pay.substr(0,num_len -6) + "," + bb_pay.substr(num_len -6,3) + "," + bb_pay.substr(num_len -2,3);
				document.frm.b_pay.value = bb_pay;
			
				bb_bonus = String(bb_bonus);
				num_len = bb_bonus.length;
				sil_len = num_len;
				bb_bonus = String(bb_bonus);
				if (bb_bonus.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_bonus = bb_bonus.substr(0,num_len -3) + "," + bb_bonus.substr(num_len -3,3);
				if (sil_len > 6) bb_bonus = bb_bonus.substr(0,num_len -6) + "," + bb_bonus.substr(num_len -6,3) + "," + bb_bonus.substr(num_len -2,3);
				document.frm.b_bonus.value = bb_bonus;
				
				deem_bonus = String(deem_bonus);
				num_len = deem_bonus.length;
				sil_len = num_len;
				deem_bonus = String(deem_bonus);
				if (deem_bonus.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) deem_bonus = deem_bonus.substr(0,num_len -3) + "," + deem_bonus.substr(num_len -3,3);
				if (sil_len > 6) deem_bonus = deem_bonus.substr(0,num_len -6) + "," + deem_bonus.substr(num_len -6,3) + "," + deem_bonus.substr(num_len -2,3);
				document.frm.b_deem_bonus.value = deem_bonus;
				
				overtime_taxno = String(overtime_taxno);
				num_len = overtime_taxno.length;
				sil_len = num_len;
				overtime_taxno = String(overtime_taxno);
				if (overtime_taxno.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) overtime_taxno = overtime_taxno.substr(0,num_len -3) + "," + overtime_taxno.substr(num_len -3,3);
				if (sil_len > 6) overtime_taxno = overtime_taxno.substr(0,num_len -6) + "," + overtime_taxno.substr(num_len -6,3) + "," + overtime_taxno.substr(num_len -2,3);
				document.frm.b_overtime_taxno.value = overtime_taxno;
				
				bb_nps = String(bb_nps);
				num_len = bb_nps.length;
				sil_len = num_len;
				bb_nps = String(bb_nps);
				if (bb_nps.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_nps = bb_nps.substr(0,num_len -3) + "," + bb_nps.substr(num_len -3,3);
				if (sil_len > 6) bb_nps = bb_nps.substr(0,num_len -6) + "," + bb_nps.substr(num_len -6,3) + "," + bb_nps.substr(num_len -2,3);
				document.frm.b_nps.value = bb_nps;
				
				bb_nhis = String(bb_nhis);
				num_len = bb_nhis.length;
				sil_len = num_len;
				bb_nhis = String(bb_nhis);
				if (bb_nhis.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_nhis = bb_nhis.substr(0,num_len -3) + "," + bb_nhis.substr(num_len -3,3);
				if (sil_len > 6) bb_nhis = bb_nhis.substr(0,num_len -6) + "," + bb_nhis.substr(num_len -6,3) + "," + bb_nhis.substr(num_len -2,3);
				document.frm.b_nhis.value = bb_nhis;
				
				bb_epi = String(bb_epi);
				num_len = bb_epi.length;
				sil_len = num_len;
				bb_epi = String(bb_epi);
				if (bb_epi.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_epi = bb_epi.substr(0,num_len -3) + "," + bb_epi.substr(num_len -3,3);
				if (sil_len > 6) bb_epi = bb_epi.substr(0,num_len -6) + "," + bb_epi.substr(num_len -6,3) + "," + bb_epi.substr(num_len -2,3);
				document.frm.b_epi.value = bb_epi;
				
				bb_longcare = String(bb_longcare);
				num_len = bb_longcare.length;
				sil_len = num_len;
				bb_longcare = String(bb_longcare);
				if (bb_longcare.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_longcare = bb_longcare.substr(0,num_len -3) + "," + bb_longcare.substr(num_len -3,3);
				if (sil_len > 6) bb_longcare = bb_longcare.substr(0,num_len -6) + "," + bb_longcare.substr(num_len -6,3) + "," + bb_longcare.substr(num_len -2,3);
				document.frm.b_longcare.value = bb_longcare;
				
				bb_income_tax = String(bb_income_tax);
				num_len = bb_income_tax.length;
				sil_len = num_len;
				bb_income_tax = String(bb_income_tax);
				if (bb_income_tax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_income_tax = bb_income_tax.substr(0,num_len -3) + "," + bb_income_tax.substr(num_len -3,3);
				if (sil_len > 6) bb_income_tax = bb_income_tax.substr(0,num_len -6) + "," + bb_income_tax.substr(num_len -6,3) + "," + bb_income_tax.substr(num_len -2,3);
				document.frm.b_income_tax.value = bb_income_tax;
				
				bb_wetax = String(bb_wetax);
				num_len = bb_wetax.length;
				sil_len = num_len;
				bb_wetax = String(bb_wetax);
				if (bb_wetax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) bb_wetax = bb_wetax.substr(0,num_len -3) + "," + bb_wetax.substr(num_len -3,3);
				if (sil_len > 6) bb_wetax = bb_wetax.substr(0,num_len -6) + "," + bb_wetax.substr(num_len -6,3) + "," + bb_wetax.substr(num_len -2,3);
				document.frm.b_wetax.value = bb_wetax;
			
			}		
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_before_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="10%" >
						<col width="15%" >
						<col width="10%" >
						<col width="15%" >
						<col width="12%" >
						<col width="13%" >
                        <col width="12%" >
						<col width="13%" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="b_emp_no" type="text" id="b_emp_no" size="10" value="<%=b_emp_no%>" readonly="true">
                      <input type="hidden" name="b_year" value="<%=b_year%>" ID="b_year">
                      <input type="hidden" name="b_seq" value="<%=b_seq%>" ID="b_seq"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="5" class="left" bgcolor="#FFFFE6">
					  <input name="b_emp_name" type="text" id="b_emp_name" size="10" value="<%=b_emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>사업자등록<br>번호</th>
                      <td class="left">
                      <input name="b_company_no" type="text" id="b_company_no" style="width:100px;text-align:left" value="<%=b_company_no%>"></td>
                      <th>근무처명</th>
                      <td class="left">
                      <input name="b_company" type="text" id="b_company" style="width:100px;text-align:left" value="<%=b_company%>"></td>
                      <th>근무시작일</th>
                      <td class="left">
					  <input name="b_from_date" type="text" value="<%=b_from_date%>" style="width:70px;text-align:center" id="datepicker" readonly="true"></td>
                      <th>근무종료일</th>
                      <td class="left">
					  <input name="b_to_date" type="text" value="<%=b_to_date%>" style="width:70px;text-align:center" id="datepicker1" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>급여</th>
					  <td class="left">
                      <input name="b_pay" type="text" id="b_pay" style="width:80px;text-align:right" value="<%=formatnumber(b_pay,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>상여</th>
					  <td class="left">
                      <input name="b_bonus" type="text" id="b_bonus" style="width:80px;text-align:right" value="<%=formatnumber(b_bonus,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>인정상여등</th>
					  <td class="left">
                      <input name="b_deem_bonus" type="text" id="b_deem_bonus" style="width:80px;text-align:right" value="<%=formatnumber(b_deem_bonus,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>비과세</th>
					  <td class="left">
                      <input name="b_overtime_taxno" type="text" id="b_overtime_taxno" style="width:80px;text-align:right" value="<%=formatnumber(b_overtime_taxno,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>국민연금</th>
					  <td class="left">
                      <input name="b_nps" type="text" id="b_nps" style="width:80px;text-align:right" value="<%=formatnumber(b_nps,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>건강보험</th>
					  <td class="left">
                      <input name="b_nhis" type="text" id="b_nhis" style="width:80px;text-align:right" value="<%=formatnumber(b_nhis,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>고용보험</th>
					  <td class="left">
                      <input name="b_epi" type="text" id="b_epi" style="width:80px;text-align:right" value="<%=formatnumber(b_epi,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>장기요양보험</th>
					  <td class="left">
                      <input name="b_longcare" type="text" id="b_longcare" style="width:80px;text-align:right" value="<%=formatnumber(b_longcare,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>(결정세액)<br>소득세</th>
					  <td class="left">
                      <input name="b_income_tax" type="text" id="b_income_tax" style="width:80px;text-align:right" value="<%=formatnumber(b_income_tax,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>(결정세액)<br>주민세</th>
					  <td colspan="4" class="left">
                      <input name="b_wetax" type="text" id="b_wetax" style="width:80px;text-align:right" value="<%=formatnumber(b_wetax,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left">※ 사업자등록번호는 전 근무지 사업자번호를 - 빼고 숫자만 입력<br>※ 근무처명은 전 근무지 법인명을 입력<br>※ 근무시작/종료일은 전 근무지 근무기간(원천징수영수증 11번 근무기간의 일자를 입력<br>※ 국민연금등은 전 근무지 원천징수 하단에보면 각각의 금액이 별도로 표시되어 있는것을 입력<br>※ 소득세등은 전 근무지 원천징수 영수증 64번 결정세액의 소득세/주민세 입력</td>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

