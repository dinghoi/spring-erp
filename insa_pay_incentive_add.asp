<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
pmg_emp_no = request("emp_no")
pmg_emp_name = request("emp_name")
pmg_yymm = request("pmg_yymm")
in_pmg_id = request("in_pmg_id")
pmg_date = request("give_date")
emp_company = request("view_condi")
view_condi = request("view_condi")
rever_year = mid(cstr(pmg_yymm),1,4) '귀속년도

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
		cost_center = rs_emp("cost_center")	  
		cost_group = rs_emp("cost_group")
   else
		emp_first_date = ""
		emp_in_date = ""
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
		cost_center = ""
		cost_group = ""
end if
rs_emp.close()

    Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
    Set rs_bnk = DbConn.Execute(SQL)
    if not rs_bnk.eof then
           bank_name = rs_bnk("bank_name")
           account_no = rs_bnk("account_no")
		   account_holder = rs_bnk("account_holder")
	   else
           bank_name = ""
		   account_no = ""
		   account_holder = ""
    end if
    rs_bnk.close()

'기본급/식대 가져오기
incom_family_cnt = 0
Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
Set Rs_year = DbConn.Execute(SQL)
if not Rs_year.eof then
		if Rs_year("incom_month_amount") = 0 or isnull(Rs_year("incom_month_amount")) then
		        incom_month_amount = Rs_year("incom_base_pay") + Rs_year("incom_overtime_pay")
		   else
		        incom_month_amount = Rs_year("incom_month_amount")
		end if
		incom_family_cnt = Rs_year("incom_family_cnt")
		incom_wife_yn = int(Rs_year("incom_wife_yn"))
		incom_age20 = Rs_year("incom_age20")
		incom_age60 = Rs_year("incom_age60")
		incom_old = Rs_year("incom_old")
		incom_disab = Rs_year("incom_disab")
		incom_go_yn = Rs_year("incom_go_yn")
   else
		incom_month_amount = 0
		incom_family_cnt = 0
		incom_wife_yn = 0
		incom_age20 = 0
		incom_age60 = 0
		incom_old = 0
		incom_disab = 0
		incom_go_yn = "여"
end if
Rs_year.close()


'if incom_family_cnt = 0 then
'    incom_family_cnt = incom_family_cnt + 1 '부양가족은 본인포함으로
'end if

incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + 1 + incom_age20 + incom_disab'본인포함 및 20세이하/장애인은 추가공제

if in_pmg_id = "2" then 
   pmg_id_name = "상여금" 
   elseif in_pmg_id = "3" then 
          pmg_id_name = "추천인인센티브" 
          elseif in_pmg_id = "4" then 
		         pmg_id_name = "연차수당" 
end if
		  
title_line = ""+ pmg_id_name +" - 자료 입력 "

if u_type = "U" then

	sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+in_pmg_id+"') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
	'Response.write sql&"<br>"
	set rs = dbconn.execute(sql)

  if not rs.eof then	
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
  	pmg_meals_pay = rs("pmg_meals_pay")
  	pmg_postage_pay = rs("pmg_postage_pay")
  	pmg_re_pay = rs("pmg_re_pay")
  	pmg_overtime_pay = rs("pmg_overtime_pay")
  	pmg_car_pay = rs("pmg_car_pay")
  	pmg_position_pay = rs("pmg_position_pay")
  	pmg_custom_pay = rs("pmg_custom_pay")
  	pmg_job_pay = rs("pmg_job_pay")
  	pmg_job_support = rs("pmg_job_support")
  	pmg_jisa_pay = rs("pmg_jisa_pay")
  	pmg_long_pay = rs("pmg_long_pay")
  	pmg_disabled_pay = rs("pmg_disabled_pay")
  	pmg_family_pay = rs("pmg_family_pay")
  	pmg_school_pay = rs("pmg_school_pay")
  	pmg_qual_pay = rs("pmg_qual_pay")
  	pmg_other_pay1 = rs("pmg_other_pay1")
  	pmg_other_pay2 = rs("pmg_other_pay2")
  	pmg_other_pay3 = rs("pmg_other_pay3")
  	pmg_tax_yes = rs("pmg_tax_yes")
  	pmg_tax_no = rs("pmg_tax_no")
  	pmg_tax_reduced = rs("pmg_tax_reduced")
  	pmg_give_tot = rs("pmg_give_total")	
  	
  	pmg_bank_name = rs("pmg_bank_name")	
  	pmg_account_no = rs("pmg_account_no")	
  	pmg_account_holder = rs("pmg_account_holder")	
  end if
	rs.close()
	
	Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '"+in_pmg_id+"') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
	'Response.write sql&"<br>"
  Set Rs_dct = DbConn.Execute(SQL)
	if not Rs_dct.eof then	
       de_nps_amt = Rs_dct("de_nps_amt")
       de_nhis_amt = Rs_dct("de_nhis_amt")
       de_epi_amt = Rs_dct("de_epi_amt")
		   de_longcare_amt = Rs_dct("de_longcare_amt")
       de_income_tax = Rs_dct("de_income_tax")
       de_wetax = Rs_dct("de_wetax")
       de_other_amt1 = Rs_dct("de_other_amt1")
		   de_special_tax = Rs_dct("de_special_tax")
       de_saving_amt = Rs_dct("de_saving_amt")
       de_sawo_amt = Rs_dct("de_sawo_amt")
       de_johab_amt = Rs_dct("de_johab_amt")
       de_hyubjo_amt = Rs_dct("de_hyubjo_amt")
       de_school_amt = Rs_dct("de_school_amt")
       de_nhis_bla_amt = Rs_dct("de_nhis_bla_amt")
       de_long_bla_amt = Rs_dct("de_long_bla_amt")	
		   'de_deduct_tot = Rs_dct("de_deduct_total")	
		   de_deduct_tot = 0
	else
		   de_deduct_tot = 0
  end if
  Rs_dct.close()	
	
	title_line = ""+ pmg_id_name +" - 자료 수정 "
end if

pay_curr_amt = pmg_give_tot - de_deduct_tot

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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=pmg_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
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
				if(document.frm.emp_no.value =="" ) {
					alert('사번을 입력하세요');
					frm.emp_no.focus();
					return false;}
				if(document.frm.pmg_date.value =="" ) {
					alert('지급일을 입력하세요');
					frm.pmg_date.focus();
					return false;}
//				if(document.frm.de_deduct_tot.value == 0 ) {
//					alert('세금계산을 하십시요');
//					frm.de_deduct_tot.focus();
//					return false;}
							
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function give_cal(txtObj){
				base_pay = parseInt(document.frm.pmg_base_pay.value.replace(/,/g,""));		

				give_tot = base_pay;
			
				base_pay = String(base_pay);
				num_len = base_pay.length;
				sil_len = num_len;
				base_pay = String(base_pay);
				if (base_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) base_pay = base_pay.substr(0,num_len -3) + "," + base_pay.substr(num_len -3,3);
				if (sil_len > 6) base_pay = base_pay.substr(0,num_len -6) + "," + base_pay.substr(num_len -6,3) + "," + base_pay.substr(num_len -2,3);
				document.frm.pmg_base_pay.value = base_pay; 
				
				give_tot = String(give_tot);
				num_len = give_tot.length;
				sil_len = num_len;
				give_tot = String(give_tot);
				if (give_tot.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) give_tot = give_tot.substr(0,num_len -3) + "," + give_tot.substr(num_len -3,3);
				if (sil_len > 6) give_tot = give_tot.substr(0,num_len -6) + "," + give_tot.substr(num_len -6,3) + "," + give_tot.substr(num_len -2,3);
				document.frm.pmg_give_tot.value = give_tot; 
				document.frm.pmg_tax_yes.value = give_tot; 
			}
			
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
 
  <script Language="JavaScript">
   function taxtax_cal() {
		if (frm.pmg_base_pay.value == 0)
		{
			alert("지급액을 입력하세요");
			frm.pmg_base_pay.focus();
			return;
		}

    var dataString = $("form").serialize();
    $.ajax({
    type: "POST",
    url : "insa_pay_incentive_taxcal.asp",
    data: dataString, //파라메터
    success: whenSuccess, //성공시 callback
    error: whenError //실패시 callback
    });
  }

    function whenSuccess(resdata) {

            var aa = resdata.split('|');
			$("div#ajaxout").html(aa[0]);
			frm.test11.value = aa[1];
			frm.de_epi_amt.value = setComma(aa[2]);
			frm.de_income_tax.value = setComma(aa[3]);
			frm.de_wetax.value = setComma(aa[4]);
			frm.de_deduct_tot.value = setComma(aa[5]);
			frm.pay_curr_amt.value = setComma(aa[6]);

    }

    function whenError(){
        alert("Error");
    }

	function setComma(str) {
      str = ""+str+"";
      var retValue = "";
      for(i=0; i<str.length; i++)
      {
        if(i > 0 && (i%3)==0) {
           retValue = str.charAt(str.length - i -1) + "," + retValue;
        } else {
           retValue = str.charAt(str.length - i -1) + retValue;
        }
      }
      return retValue;
		}        
        
 </script>               
        
        
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_incentive_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">사번</th>
								<td class="left">
                                <input name="emp_no" type="text" value="<%=pmg_emp_no%>" style="width:90px" readonly="true"></td>
								<th >성명</th>
								<td class="left" >
                                <input name="pmg_emp_name" type="text" value="<%=pmg_emp_name%>" style="width:90px" readonly="true"></td>
							</tr>
                           	<tr>
								<th class="first">직급</th>
								<td class="left"><input name="pmg_grade" type="text" value="<%=pmg_grade%>" style="width:90px" readonly="true"></td>
                                <th >직책</th>
								<td class="left" ><input name="pmg_position" type="text" value="<%=pmg_position%>" style="width:90px" readonly="true"></td>
							</tr>    
                            <tr>
								<th class="first">귀속년월</th>
								<td class="left" ><input name="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:70px" readonly="true"></td>
                                <th >지급일</th>
                                <td class="left"><input name="pmg_date" type="text" value="<%=pmg_date%>" style="width:70px" id="datepicker"></td>
							</tr>             
							<tr>
								<th class="first">소속</th>
								<td class="left"><%=pmg_company%>&nbsp;&nbsp;<%=pmg_org_name%>(<%=pmg_org_code%>)&nbsp;</td>
								<th>계좌번호</th>
								<td class="left"><%=account_no%>(<%=bank_name%>-<%=account_holder%>)&nbsp;</td>
							</tr>
                            <tr>
								<th colspan="2" class="first" style="background:#F5FFFA">지급항목</th>
								<th colspan="2" class="first" style="background:#F8F8FF">공제항목</th>
                                <input type="hidden" name="rever_year" value="<%=rever_year%>" ID="Hidden1">
                                <input type="hidden" name="incom_family_cnt" value="<%=incom_family_cnt%>" ID="Hidden1">
                                <input type="hidden" name="in_pmg_id" value="<%=in_pmg_id%>" ID="Hidden1">
                                <input type="hidden" name="incom_go_yn" value="<%=incom_go_yn%>" ID="Hidden1">
							</tr>  
							<tr>
								<%
								  if in_pmg_id = "2" then %>
								  <th class="first" style="background:#F5FFFA">상여금</th>
                                <%   elseif in_pmg_id = "3" then %>
                                     <th class="first" style="background:#F5FFFA">추천인<br>인센티브</th>
                                <%          elseif in_pmg_id = "4" then %>
                                            <th class="first" style="background:#F5FFFA">연차수당</th>
                                <% end if %>
								<td class="left">
                    <input name="pmg_base_pay" type="text" value="<%=formatnumber(pmg_base_pay,0)%>" style="width:100px;text-align:right" onKeyUp="give_cal(this);"></td>
								<th style="background:#F8F8FF">고용보험</th>
                <td class="left">
								<input name="de_epi_amt" type="text" value="<%=formatnumber(de_epi_amt,0)%>" style="width:100px;text-align:right"></td>
							</tr>
              <tr>
								<th class="first" style="background:#F5FFFA">&nbsp;</th>
								<td class="left">&nbsp;<input name="ajaxout" type="hidden" id="ajaxout" size="14" value="<%=ajaxout%>"></td>                
								<th style="background:#F8F8FF">소득세</th>
                <td class="left">
								<input name="de_income_tax" type="text" value="<%=formatnumber(de_income_tax,0)%>" style="width:100px;text-align:right"></td>
							</tr>
              <tr>
								<th class="first" style="background:#F5FFFA">&nbsp;</th>
								<td class="left">
                    <input name="test11" type="hidden" value="<%=test11%>" style="width:100px;text-align:center">
                </td>
								<th style="background:#F8F8FF">지방소득세</th>
                <td class="left">
								<input name="de_wetax" type="text" value="<%=formatnumber(de_wetax,0)%>" style="width:100px;text-align:right"></td>
							</tr>   
              <tr>
								<th class="first" style="background:#F5FFFA">&nbsp;</th>
								<td class="left">&nbsp;</td>
                <input name="pmg_tax_reduced" type="hidden" value="<%=formatnumber(pmg_tax_reduced,0)%>" style="width:100px;text-align:right"></td>
								<th style="background:#F8F8FF">공제액 계</th>
                <td class="left">
								<input name="de_deduct_tot" type="text" value="<%=formatnumber(de_deduct_tot,0)%>" style="width:100px;text-align:right"></td>
							</tr>    
              <tr>
								<th class="first" style="background:#F5FFFA">지급액 계</th>
								<td class="left">
                    <input name="pmg_give_tot" type="text" value="<%=formatnumber(pmg_give_tot,0)%>" style="width:100px;text-align:right" readonly="true">
                    <input name="pmg_tax_yes" type="hidden" value="<%=formatnumber(pmg_give_tot,0)%>" style="width:100px;text-align:right">
                </td>
								<th style="background:#F8F8FF">차인지급액</th>
                <td class="left">
								    <input name="pay_curr_amt" type="text" value="<%=formatnumber(pay_curr_amt,0)%>" style="width:100px;text-align:right">
                    <a href="#" onClick="javascript:taxtax_cal();" class="btn-gray2">세금계산</a>
                </td>
							</tr>              
            </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="pmg_company" value="<%=pmg_company%>" ID="Hidden1">
                <input type="hidden" name="pmg_bonbu" value="<%=pmg_bonbu%>" ID="Hidden1">
                <input type="hidden" name="pmg_saupbu" value="<%=pmg_saupbu%>" ID="Hidden1">
                <input type="hidden" name="pmg_team" value="<%=pmg_team%>" ID="Hidden1">
                <input type="hidden" name="pmg_reside_place" value="<%=pmg_reside_place%>" ID="Hidden1">
                <input type="hidden" name="pmg_reside_company" value="<%=pmg_reside_company%>" ID="Hidden1">
                <input type="hidden" name="emp_in_date" value="<%=emp_in_date%>" ID="Hidden1">
                <input type="hidden" name="cost_group" value="<%=cost_group%>" ID="Hidden1">
                <input type="hidden" name="cost_center" value="<%=cost_center%>" ID="Hidden1">
                <input type="hidden" name="pmg_org_name" value="<%=pmg_org_name%>" ID="Hidden1">
                <input type="hidden" name="pmg_org_code" value="<%=pmg_org_code%>" ID="Hidden1">
                <input type="hidden" name="pmg_emp_type" value="<%=pmg_emp_type%>" ID="Hidden1">
                <input type="hidden" name="pmg_bank_name" value="<%=bank_name%>" ID="Hidden1">
                <input type="hidden" name="pmg_account_no" value="<%=account_no%>" ID="Hidden1">
                <input type="hidden" name="pmg_account_holder" value="<%=account_holder%>" ID="Hidden1">
			</form> 
		</div>				
	</body>
</html>

