<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
inc_yyyy = Request("inc_yyyy")

incom_base_pay = 0
incom_overtime_pay = 0
incom_meals_pay = 0
incom_severance_pay = 0
incom_total_pay = 0
incom_month_amount = 0
incom_nps_amount = 0
incom_nhis_amount = 0
incom_nps = 0
incom_nhis = 0
incom_go_yn = "여"
incom_san_yn = "여"
incom_long_yn = "여"
incom_incom_yn = "부"
incom_family_cnt = 0
incom_wife_yn = "0"
incom_age20 = 0
incom_age60 = 0
incom_old = 0
incom_disab = 0
incom_woman = "0"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 국민연금 표준월액 등록 "

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
Rs_emp.Open Sql, Dbconn, 1

	incom_in_date = Rs_emp("emp_in_date")
	incom_grade = Rs_emp("emp_grade")
	incom_emp_type = Rs_emp("emp_type")
	if Rs_emp("emp_pay_type") = "1" then 
	      incom_pay_type = "근로소득"
	   else
	      incom_pay_type = "사업소득"	
    end if  
	incom_company = Rs_emp("emp_company")
	incom_org_code = Rs_emp("emp_org_code")
	incom_org_name = Rs_emp("emp_org_name")

incom_year = curr_year

'국민연금 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&inc_yyyy&"' and insu_id = '5501' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nps_emp = formatnumber(rs_ins("emp_rate"),3)
		nps_com = formatnumber(rs_ins("com_rate"),3)
		nps_from = rs_ins("from_amt")
		nps_to = rs_ins("to_amt")
   else
		nps_emp = 0
		nps_com = 0
		nps_from = 0
		nps_to = 0
end if
rs_ins.close()

'건강보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&inc_yyyy&"' and insu_id = '5502' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nhis_emp = formatnumber(rs_ins("emp_rate"),3)
		nhis_com = formatnumber(rs_ins("com_rate"),3)
		nhis_from = rs_ins("from_amt")
		nhis_to = rs_ins("to_amt")
   else
		nhis_emp = 0  
		nhis_com = 0
		nhis_from = 0
		his_to = 0
end if
rs_ins.close()

'if u_type = "U" then

	Sql="select * from pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&inc_yyyy&"'"
	Set rs=DbConn.Execute(Sql)
  if not rs.eof then
    u_type = "U"
    incom_year = rs("incom_year")
	incom_in_date = rs("incom_in_date")
	incom_grade = rs("incom_grade")
	incom_emp_type = rs("incom_emp_type")
	
	if rs("incom_pay_type") = "1" then 
	      incom_pay_type = "근로소득"
	   else
	      incom_pay_type = "사업소득"	
    end if  
	incom_company = rs("incom_company")
	incom_org_code = rs("incom_org_code")
	incom_org_name = rs("incom_org_name")
	
	incom_base_pay = rs("incom_base_pay")
    incom_overtime_pay = rs("incom_overtime_pay")
    incom_meals_pay = rs("incom_meals_pay")
    incom_severance_pay = rs("incom_severance_pay")
	incom_total_pay = rs("incom_total_pay")
	incom_month_amount = rs("incom_month_amount")
	incom_nps_amount = rs("incom_nps_amount")
	incom_nhis_amount = rs("incom_nhis_amount")
	incom_family_cnt = rs("incom_family_cnt")
	incom_nps = rs("incom_nps")
    incom_nhis = rs("incom_nhis")
    incom_go_yn = rs("incom_go_yn")
    incom_san_yn = rs("incom_san_yn")
    incom_long_yn = rs("incom_long_yn")
    incom_incom_yn = rs("incom_incom_yn")
    incom_wife_yn = rs("incom_wife_yn")
    incom_age20 = rs("incom_age20")
    incom_age60 = rs("incom_age60")
    incom_old = rs("incom_old")
    incom_disab = rs("incom_disab")
    incom_woman = rs("incom_woman")
	
	if isnull(incom_go_yn) or incom_go_yn = "" then
	      incom_go_yn = "여"
	end if
	if isnull(incom_san_yn) or incom_san_yn = "" then
	      incom_san_yn = "여"
	end if
	if isnull(incom_long_yn) or incom_long_yn = "" then
	      incom_long_yn = "여"
	end if
	if isnull(incom_incom_yn) or incom_incom_yn = "" then
	      incom_incom_yn = "부"
	end if
	if isnull(incom_wife_yn) or incom_wife_yn = "" then
	      incom_wife_yn = "0"
	end if
	if isnull(incom_woman) or incom_woman = "" then
	      incom_woman = "0"
	end if
	
	rs.close()

	title_line = " 국민연금 표준월액 변경 "
  else
    u_type = ""	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
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
				if(document.frm.incom_nps_amount.value =="") {
					alert('국민연금표준월액을 입력하세요');
					frm.incom_nps_amount.focus();
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
				nps_amount = parseInt(document.frm.incom_nps_amount.value.replace(/,/g,""));	
				
				if (nps_amount > 4210000) nps_amount = 4210000;
				
				e_nps = parseFloat((document.frm.nps_emp.value),3);
				
                nps_amt = nps_amount * (e_nps / 100);
				nps_amt = parseInt(nps_amt);
				nps_amt = (parseInt(nps_amt / 10)) * 10;
			
				nps_amount = String(nps_amount);
				num_len = nps_amount.length;
				sil_len = num_len;
				nps_amount = String(nps_amount);
				if (nps_amount.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nps_amount = nps_amount.substr(0,num_len -3) + "," + nps_amount.substr(num_len -3,3);
				if (sil_len > 6) nps_amount = nps_amount.substr(0,num_len -6) + "," + nps_amount.substr(num_len -6,3) + "," + nps_amount.substr(num_len -2,3);
				document.frm.incom_nps_amount.value = nps_amount;
				
				nps_amt = String(nps_amt);
				num_len = nps_amt.length;
				sil_len = num_len;
				nps_amt = String(nps_amt);
				if (nps_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nps_amt = nps_amt.substr(0,num_len -3) + "," + nps_amt.substr(num_len -3,3);
				if (sil_len > 6) nps_amt = nps_amt.substr(0,num_len -6) + "," + nps_amt.substr(num_len -6,3) + "," + nps_amt.substr(num_len -2,3);
				document.frm.incom_nps.value = nps_amt;
		
			}							
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_year_nps_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="10%">
						<col width="15%">
						<col width="10%">
						<col width="15%">
						<col width="10%">
						<col width="15%">
                        <col width="10%">
						<col width="15%">
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6"><%=emp_no%>
					  <input name="emp_no" type="hidden" id="emp_no" size="14" value="<%=emp_no%>" readonly="true"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td class="left" bgcolor="#FFFFE6"><%=emp_name%>
					  <input name="emp_name" type="hidden" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                      <th style="background:#FFFFE6">입사일</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6"><%=incom_in_date%>
					  <input name="incom_in_date" type="hidden" id="incom_in_date" size="14" value="<%=incom_in_date%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">직급</th>
                      <td class="left" bgcolor="#FFFFE6"><%=incom_grade%>
					  <input name="incom_grade" type="hidden" id="incom_grade" size="14" value="<%=incom_grade%>" readonly="true"></td>
                      <th style="background:#FFFFE6">직원구분</th>
                      <td class="left" bgcolor="#FFFFE6"><%=incom_emp_type%>
					  <input name="incom_emp_type" type="hidden" id="incom_emp_type" size="14" value="<%=incom_emp_type%>" readonly="true"></td>
                      <th style="background:#FFFFE6">소득구분</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6"><%=incom_pay_type%>
					  <input name="incom_pay_type" type="hidden" id="incom_pay_type" size="14" value="<%=incom_pay_type%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">회사</th>
                      <td class="left" bgcolor="#FFFFE6"><%=incom_company%>
					  <input name="incom_company" type="hidden" id="incom_company" size="19" value="<%=incom_company%>" readonly="true"></td>
                      <th style="background:#FFFFE6">소속</th>
                      <td colspan="5" class="left" bgcolor="#FFFFE6"><%=incom_org_name%> - <%=incom_org_code%>
					  <input name="incom_org_name" type="hidden" id="incom_org_name" size="19" value="<%=incom_org_name%>" readonly="true">
                      -
					  <input name="incom_org_code" type="hidden" id="incom_org_code" size="6" value="<%=incom_org_code%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>귀속년도</th>
                      <td colspan="7" class="left"><%=incom_year%>
					  <input name="incom_year" type="hidden" id="incom_year" size="7" value="<%=incom_year%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#F5FFFA">국민연금<br>표준월액</th>
                      <td class="left">
                      <input name="incom_nps_amount" type="text" id="incom_nps_amount" style="width:80px;text-align:right" value="<%=formatnumber(incom_nps_amount,0)%>" onKeyUp="num_chk(this);"></td>
                      <th style="background:#F5FFFA">국민연금</th>
                      <td colspan="5" class="left">
                      <input name="incom_nps" type="text" id="incom_nps" style="width:80px;text-align:right" value="<%=formatnumber(incom_nps,0)%>" readonly="true"></td>
			    	</tr>
                    <tr>
                      <th style="background:#F5FFFA">고용보험<br>적용</th>
                      <td class="left">
                      <input type="radio" name="incom_go_yn" value= "여" <% if incom_go_yn = "여" then %>checked<% end if %>>여 
              		  <input name="incom_go_yn" type="radio" value= "부" <% if incom_go_yn = "부" then %>checked<% end if %>>부
                      </td>
                      <th style="background:#F5FFFA">산재보험<br>적용</th>
                      <td class="left">
                      <input type="radio" name="incom_san_yn" value= "여" <% if incom_san_yn = "여" then %>checked<% end if %>>여 
              		  <input name="incom_san_yn" type="radio" value= "부" <% if incom_san_yn = "부" then %>checked<% end if %>>부
                      </td>
                      <th style="background:#F5FFFA">장기요양<br>보험적용</th>
                      <td class="left">
                      <input type="radio" name="incom_long_yn" value= "여" <% if incom_long_yn = "여" then %>checked<% end if %>>여 
              		  <input name="incom_long_yn" type="radio" value= "부" <% if incom_long_yn = "부" then %>checked<% end if %>>부
                      </td>
                      <th style="background:#F5FFFA">청년<br>소득세감면</th>
                      <td class="left">
                      <input type="radio" name="incom_incom_yn" value= "여" <% if incom_incom_yn = "여" then %>checked<% end if %>>여 
              		  <input name="incom_incom_yn" type="radio" value= "부" <% if incom_incom_yn = "부" then %>checked<% end if %>>부
                      </td>
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
                <input type="hidden" name="nps_emp" value="<%=formatnumber(nps_emp,3)%>" ID="Hidden1">
                <input type="hidden" name="nps_com" value="<%=formatnumber(nps_com,3)%>" ID="Hidden1">
                <input type="hidden" name="nhis_emp" value="<%=formatnumber(nhis_emp,3)%>" ID="Hidden1">
                <input type="hidden" name="nhis_com" value="<%=formatnumber(nhis_com,3)%>" ID="Hidden1">
                <input type="hidden" name="nps_from" value="<%=nps_from%>" ID="Hidden1">
                <input type="hidden" name="nps_to" value="<%=nps_to%>" ID="Hidden1">
                <input type="hidden" name="nhis_from" value="<%=nhis_from%>" ID="Hidden1">
                <input type="hidden" name="nhis_to" value="<%=nhis_to%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

