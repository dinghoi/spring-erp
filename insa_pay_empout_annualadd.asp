<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

emp_no = request("emp_no")
emp_name = request("emp_name")
view_condi = request("view_condi")
pmg_yymm = request("pmg_yymm")
u_type = request("u_type")
rever_year = mid(cstr(pmg_yymm),1,4) '귀속년도
curr_date = mid(cstr(now()),1,10)
pmg_date = curr_date

	pmg_org_code = ""
	pmg_org_name = ""
	pmg_emp_type = ""
	pmg_grade = ""
	pmg_position = ""
	
	pmg_base_pay = 0
	pmg_meals_pay = 0
	pmg_postage_pay = 0
	pmg_re_pay = 0
	pmg_overtime_pay = 0
	pmg_car_pay = 0
	pmg_position_pay = 0
	pmg_custom_pay = 0
	pmg_job_pay = 0
	pmg_job_support = 0
	pmg_jisa_pay = 0
	pmg_long_pay = 0
	pmg_disabled_pay = 0
	pmg_family_pay = 0
	pmg_school_pay = 0
	pmg_qual_pay = 0
	pmg_other_pay1 = 0
	pmg_other_pay2 = 0
	pmg_other_pay3 = 0
	pmg_tax_yes = 0
	pmg_tax_no = 0
	pmg_tax_reduced = 0
	
    de_nps_amt = 0
    de_nhis_amt = 0
    de_epi_amt = 0
	de_longcare_amt = 0
    de_income_tax = 0
    de_wetax = 0
    de_special_tax = 0
    de_saving_amt = 0
    de_sawo_amt = 0
    de_johab_amt = 0
    de_hyubjo_amt = 0
    de_school_amt = 0
    de_nhis_bla_amt = 0
    de_long_bla_amt = 0
	
    pay_curr_amt = 0
	pmg_give_tot = 0
	de_deduct_tot = 0

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_first_date = rs_emp("emp_first_date")
		emp_in_date = rs_emp("emp_in_date")
		emp_end_date = rs_emp("emp_end_date")
		pmg_emp_type = rs_emp("emp_type")
		pmg_grade = rs_emp("emp_grade")
		pmg_position = rs_emp("emp_position")
		pmg_company = rs_emp("emp_company")
		pmg_bonbu = rs_emp("emp_bonbu")
		pmg_saupbu = rs_emp("emp_saupbu")
		pmg_team = rs_emp("emp_team")
		pmg_org_code = rs_emp("emp_org_code")
		pmg_org_name = rs_emp("emp_org_name")
		pmg_reside_place = rs_emp("emp_reside_place")
		pmg_reside_company = rs_emp("emp_reside_company")
		if rs_emp("emp_yuncha_date") = "1900-01-01" or isNull(rs_emp("emp_yuncha_date")) then
                emp_yuncha_date = rs_emp("emp_in_date")
           else 
                emp_yuncha_date = rs_emp("emp_yuncha_date")
        end if
   else
		emp_first_date = ""
		emp_in_date = ""
		emp_end_date = ""
		emp_yuncha_date = ""
		pmg_emp_type = ""
		pmg_grade = ""
		pmg_position = ""
		pmg_company = ""
		pmg_bonbu = ""
		pmg_saupbu = ""
		pmg_team = ""
		pmg_org_code = ""
		pmg_org_name = ""
		pmg_reside_place = ""
		pmg_reside_company = ""
end if
rs_emp.close()

' 근속년수
target_date = emp_end_date + 1
year_cnt = datediff("yyyy", emp_yuncha_date, target_date)
					  
' 연차일수
target_date = emp_end_date
if (datediff("d", emp_yuncha_date, target_date) + 1) / 365 < 1 then
         yun_day = datediff("m", emp_yuncha_date, target_date) 
     else
	     yun_day = round((((datediff("d", emp_yuncha_date, target_date) + 1) / 365) / 2),0) + 14
end if
							  
' 누적연차수
if datediff("yyyy", emp_yuncha_date, target_date) mod 2 = 1 then
          tot_yun = round(((year_cnt ^ 2 + 58 * year_cnt - 0) / 4),0)
	 else
          tot_yun = year_cnt / 2 * (year_cnt / 2 + 1) + 14 * year_cnt
end if
							  
mon_cnt = datediff("m", emp_yuncha_date, target_date) 


'고용보험(실업) 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5503' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	epi_emp = formatnumber(rs_ins("emp_rate"),3)
		epi_com = formatnumber(rs_ins("com_rate"),3)
   else
		epi_emp = 0  
		epi_com = 0
end if
rs_ins.close()

'기본급/식대 가져오기
Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
Set Rs_year = DbConn.Execute(SQL)
if not Rs_year.eof then
    	pmg_base_pay = Rs_year("incom_base_pay")
		pmg_meals_pay = Rs_year("incom_meals_pay")
		pmg_overtime_pay = Rs_year("incom_overtime_pay")
   else
		pmg_base_pay = 0  
		pmg_meals_pay = 0
		pmg_overtime_pay = 0
end if
Rs_year.close()

month_pay = pmg_base_pay + pmg_meals_pay + pmg_overtime_pay

title_line = " 퇴직자 연차정산 등록 "

if u_type = "U" then

	sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '4') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
	set rs = dbconn.execute(sql)

    pmg_yymm = rs("pmg_yymm")
	pmg_emp_no = rs("pmg_emp_no")
    pmg_company = rs("pmg_company")
	pmg_date = rs("pmg_date")
	pmg_emp_name = rs("pmg_emp_name")
	pmg_org_code = rs("pmg_org_code")
	pmg_org_name = rs("pmg_org_name")
	pmg_emp_type = rs("pmg_emp_type")
	pmg_grade = rs("pmg_grade")
	pmg_position = rs("pmg_position")
	
	pmg_base_pay = rs("pmg_base_pay")
	pmg_give_tot = rs("pmg_give_total")	

	rs.close()
	
	Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '4') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
	if not Rs_dct.eof then	
           de_epi_amt = Rs_dct("de_epi_amt")
		   de_deduct_tot = Rs_dct("de_deduct_total")
	   else
		   de_deduct_tot = 0
		   de_epi_amt = 0
    end if
    Rs_dct.close()	
	pay_curr_amt = pmg_give_tot - de_deduct_tot
	de_deduct_tot = 0

	title_line = " 퇴직자 연차정산 변경 "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=pmg_date%>" );
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
				if(document.frm.family_birthday.value =="") {
					alert('생년월일을 입력하세요');
					frm.family_birthday.focus();
					return false;}
				if(document.frm.family_rel =="") {
					alert('관계항목을 선택하세요');
					frm.family_rel.focus();
					return false;}
				if(document.frm.family_name.value =="") {
					alert('가족성명을 입력하세요');
					frm.family_name.focus();
					return false;}
				if(document.frm.family_tel_no1.value =="") {
					alert('전화번호를 입력하세요');
					frm.family_tel_no1.focus();
					return false;}
				if(document.frm.family_tel_no2.value =="") {
					alert('전화번호를 입력하세요');
					frm.family_tel_no2.focus();
					return false;}
				if(document.frm.family_support_yn.value =="") {
					alert('부양가족여부를 입력하세요');
					frm.family_support_yn.focus();
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
				t_yun = parseInt(document.frm.tot_yun.value.replace(/,/g,""));	
				u_yun = parseInt(document.frm.use_yun.value.replace(/,/g,""));	
				j_yun = parseInt(document.frm.just_yun.value.replace(/,/g,""));	
				m_pay = parseInt(document.frm.month_pay.value.replace(/,/g,""));
				e_epi = parseFloat((document.frm.epi_emp.value),3);
				
				n_yun = parseInt(t_yun - u_yun - j_yun);
				
				document.frm.jan_yun.value = n_yun; 
				
				y_pay = parseInt(m_pay / 209 * 8 * n_yun);
				
				epi_amt = y_pay * (e_epi /100);
				epi_amt = parseInt(epi_amt);
				e_amt = (parseInt(epi_amt / 10)) * 10;
				
				c_yun_pay = y_pay - e_amt
			
				y_pay = String(y_pay);
				num_len = y_pay.length;
				sil_len = num_len;
				y_pay = String(y_pay);
				if (y_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) y_pay = y_pay.substr(0,num_len -3) + "," + y_pay.substr(num_len -3,3);
				if (sil_len > 6) y_pay = y_pay.substr(0,num_len -6) + "," + y_pay.substr(num_len -6,3) + "," + y_pay.substr(num_len -2,3);
				document.frm.yun_pay.value = y_pay; 
				
				e_amt = String(e_amt);
				num_len = e_amt.length;
				sil_len = num_len;
				e_amt = String(e_amt);
				if (e_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) e_amt = e_amt.substr(0,num_len -3) + "," + e_amt.substr(num_len -3,3);
				if (sil_len > 6) e_amt = e_amt.substr(0,num_len -6) + "," + e_amt.substr(num_len -6,3) + "," + e_amt.substr(num_len -2,3);
				document.frm.epi_amt.value = e_amt; 
				
				c_yun_pay = String(c_yun_pay);
				num_len = c_yun_pay.length;
				sil_len = num_len;
				c_yun_pay = String(c_yun_pay);
				if (c_yun_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) c_yun_pay = c_yun_pay.substr(0,num_len -3) + "," + c_yun_pay.substr(num_len -3,3);
				if (sil_len > 6) c_yun_pay = c_yun_pay.substr(0,num_len -6) + "," + c_yun_pay.substr(num_len -6,3) + "," + c_yun_pay.substr(num_len -2,3);
				document.frm.curr_pay.value = c_yun_pay; 
				
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_empout_annual_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="emp_no" type="text" id="emp_no" size="7" value="<%=emp_no%>" readonly="true">
                      <th style="background:#FFFFE6">성명</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="10" value="<%=emp_name%>" readonly="true"></td>
                      <th style="background:#FFFFE6">입사일</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="emp_in_date" type="text" id="emp_in_date" size="10" value="<%=emp_in_date%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">퇴사일</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="emp_end_date" type="text" id="emp_end_date" size="10" value="<%=emp_end_date%>" readonly="true">
                      <th style="background:#FFFFE6">연차기산일</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="emp_yuncha_date" type="text" id="emp_yuncha_date" size="10" value="<%=emp_yuncha_date%>" readonly="true"></td>
                      <th style="background:#FFFFE6">근속년수</th>
                      <td class="left" bgcolor="#FFFFE6"><%=year_cnt%>년&nbsp;(총:&nbsp;<%=mon_cnt%>개월)</td>
                    </tr>
                    <tr> 
                      <th>귀속(지급)<br>년월</th>
                      <td class="left">
					  <input name="pmg_yymm" type="text" id="pmg_yymm" size="6" value="<%=pmg_yymm%>" readonly="true"></td>
                      <th>지급일</th>
                      <td colspan="3" class="left">
					  <input name="pmg_date" type="text" value="<%=pmg_date%>" style="width:70px;text-align:center" id="datepicker" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">누적연차</th>
                      <td colspan="5" class="left" bgcolor="#FFFFE6">
					  <input name="tot_yun" type="text" id="tot_yun" size="5" value="<%=tot_yun%>" style="width:50px;text-align:right"readonly="true"></td>
                    </tr>
                    <tr>
                      <th>사용연차</th>
                      <td class="left">
                      <input name="use_yun" type="text" id="use_yun" style="width:50px;text-align:right" value="<%=formatnumber(use_yun,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>정산연차</th>
                      <td class="left">
                      <input name="just_yun" type="text" id="just_yun" style="width:50px;text-align:right" value="<%=formatnumber(just_yun,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>잔여연차</th>
                      <td class="left">
                      <input name="jan_yun" type="text" id="jan_yun" style="width:50px;text-align:right" readonly="true" value="<%=formatnumber(jan_yun,0)%>"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">월급여</th> 
                      <td colspan="5" class="left" bgcolor="#FFFFE6">
					  <input name="month_pay" type="text" id="month_pay" style="width:90px;text-align:right" value="<%=formatnumber(month_pay,0)%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>연차수당</th>
                      <td class="left">
                      <input name="yun_pay" type="text" id="yun_pay" style="width:90px;text-align:right" readonly="true" value="<%=formatnumber(yun_pay,0)%>"></td>
                      <th>고용보험</th>
                      <td class="left">
                      <input name="epi_amt" type="text" id="epi_amt" style="width:90px;text-align:right" readonly="true" value="<%=formatnumber(epi_amt,0)%>"></td>
                      <th>실지급액</th>
                      <td class="left">
                      <input name="curr_pay" type="text" id="curr_pay" style="width:80px;text-align:right" readonly="true" value="<%=formatnumber(curr_pay,0)%>"></td>
                    </tr>
                    <tr>
                      <th>비고</th>
                      <td colspan="5" class="left">
                      <textarea name="pmg_comment" rows="1" id="textarea"><%=pmg_comment%></textarea></td>
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
                <input type="hidden" name="epi_emp" value="<%=formatnumber(epi_emp,3)%>" ID="Hidden1">
                <input type="hidden" name="epi_com" value="<%=formatnumber(epi_com,3)%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

