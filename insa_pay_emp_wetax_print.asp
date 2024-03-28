<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_pay_emp_wetax_print.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = request("view_condi")
pmg_yymm = request("pmg_yymm")

'tax_date = cstr(mid(dateadd("m",+1,now()),1,4)) + "-" + cstr(mid(dateadd("m",+1,now()),6,2)) + "-" + "10"	
tax_date = curr_year + "-" + curr_month + "-" + "10"

	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	
	
	tax_meals_no = 0	
	tax_car_no = 0	
	tax_meals_yes = 0	
	tax_car_yes = 0	
	

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	  pay_count = pay_count + 1
							  
	  pmg_date = rs("pmg_date")
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

	  meals_pay = pmg_meals_pay
	  car_pay = pmg_car_pay
	  meals_tax_pay = 0
	  meals_taxno_pay = 0
	  car_tax_pay = 0
	  car_taxno_pay = 0
	  
	  if  meals_pay > 100000 then
	         meals_tax_pay = meals_pay - 100000
	         tax_meals_yes = tax_meals_yes + (meals_pay - 100000)
			 meals_taxno_pay = 100000
			 tax_meals_no= tax_meals_no + 100000
		  else	 
		     meals_taxno_pay = meals_pay
			 tax_meals_no= tax_meals_no + meals_pay
	  end if
  	  if car_pay > 200000 then
	         car_tax_pay = car_pay - 200000
			 tax_car_yes = tax_car_yes + (car_pay - 200000)
			 car_taxno_pay = 200000
			 tax_car_no =  tax_car_no + 200000
		 else
			 tax_car_no =  tax_car_no + car_pay
			 car_taxno_pay = car_pay
	  end if
	  
	  pmg_tax_yes = 0
	  pmg_tax_no = 0
	  
	  pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	  pmg_tax_no = meals_taxno_pay + car_taxno_pay
	  
	  sum_tax_yes = sum_tax_yes + pmg_tax_yes
	  sum_tax_no = sum_tax_no + pmg_tax_no

	rs.movenext()
loop
rs.close()

sum_give_tot = sum_tax_yes + sum_tax_no

month_person_pay = int(sum_tax_yes / pay_count) '신고월 월적용급여액
deduct_14 = month_person_pay * (pay_count - pay_count) '공제액
income_pay15 = sum_tax_yes - deduct_14 '산출과표
income_tax16 = int(income_pay15 * (0.5 / 100)) '산출세액
add_tax1 = 0
add_tax2 = 0
add_tax17 = 0
tax_hap = income_tax16 + add_tax17

' 금액을 한글로 변환....
'amt = "21345000"
amt = tax_hap
Dim unit1(10)
Dim unit2(2)
Dim unit3(2)

unit1(0) = ""
unit1(1) = "일"
unit1(2) = "이"
unit1(3) = "삼"
unit1(4) = "사"
unit1(5) = "오"
unit1(6) = "육"
unit1(7) = "칠"
unit1(8) = "팔"
unit1(9) = "구"

unit2(0) = "십"
unit2(1) = "백"
unit2(2) = "천"

unit3(0) = "만"
unit3(1) = "억"
unit3(2) = "조"
 
vamt = Replace(amt, ",", "")
xchk = IsNumeric(vamt)

If xchk = True Then
    total = Len(CStr(CDbl(amt)))
    vamt = CDbl(amt)
    rt_amt = ""
    For i = 1 To total
        num = Mid(vamt, i, 1)
        temp1 = (total - i) + 1
        rt_amt = rt_amt & unit1(num)
 
        If num <> 0 And i <> total Then
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 0 Then rt_amt = rt_amt & unit2(2)
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 3 Then rt_amt = rt_amt & unit2(1)
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 2 Then rt_amt = rt_amt & unit2(0)
        End If
 
        If temp1 = 5 And Right(rt_amt, 1) <> unit3(2) And Right(rt_amt, 1) <> unit3(1) Then rt_amt = rt_amt & unit3(0)
        If temp1 = 9 And Right(rt_amt, 1) <> unit3(2) Then rt_amt = rt_amt & unit3(1)
        If temp1 = 13 Then rt_amt = rt_amt & unit3(2)
 
    Next
 
    rt_amt = rt_amt & "원"
 
    'msgbox
    'response.write  "input : " & amt & vbCr & "output : " & rt_amt
End If

jiyun_day = 0


curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = " 종업원할사업소세(지방세) - 납부서 "

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
				return "5 1";
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
				if (document.frm.view_condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			

			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_emp_wetax_print.asp" method="post" name="frm">
                <h3 class="stit">*종업원분 주민세&nbsp;납부서</h3>
				<div class="gView">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tableList">
				        <tr>
                            <td width="50%" class="left">&nbsp;&nbsp;&nbsp;&nbsp;귀속년월:&nbsp;<%=mid(pmg_yymm,1,4)%>년&nbsp;<%=mid(pmg_yymm,5,2)%>월분</td>
                            <td width="50%" class="right">영수일(납부기한):&nbsp;<%=tax_date%></td>
                        </tr>
                    </table>
					<table width="100%"  cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%">
                            <col width="20%">
                            <col width="20%">
                            <col width="20%">
                            <col width="20%">
						</colgroup>
						<thead>
							<tr>
				                <th rowspan="2" class="first" scope="col">구분</th>
                                <th rowspan="2" scope="col">사업소인원</th>
				                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">과세표준액</th>
			                </tr>
                            <tr>
							    <th scope="col" style=" border-left:1px solid #e3e3e3;">과세제외급여</th>
								<th scope="col">과세급여</th>  
								<th scope="col">총지급급여액</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first" style="background:#f8f8f8;">종업원분</td>
                                <td class="right"><%=formatnumber(pay_count,0)%>&nbsp;인&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_tax_no,0)%>&nbsp;원&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_tax_yes,0)%>&nbsp;원&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_give_tot,0)%>&nbsp;원&nbsp;</td>
							</tr>
						</tbody>
					</table>
                    <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="30%" >
                            <col width="40%" >
                            <col width="30%" >
						</colgroup>
						<thead>
                            <tr>
							    <th colspan="3" class="first" scope="col">신고 납부(납입)세액</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                                <td class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">지방소득세(종업원분)</td>  
                                <td class="right"><%=rt_amt%>&nbsp;</td>
                                <td class="right"><%=formatnumber(tax_hap,0)%>&nbsp;원&nbsp;</td>
							</tr>
						</tbody>
					</table>
                    <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
						</colgroup>
						<thead>
                            <tr>
							    <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">가산세</th>
							</tr>
                            <tr>
							    <th class="first" scope="col">당초납부기한</th>
                                <th scope="col">납부지연일수</th>
                                <th scope="col">납부불성실가산세</th>
                                <th scope="col">신고불성실가산세</th>
							</tr>
                        </thead>
						<tbody>
                            <tr>
							    <td class="first" scope="col"><%=tax_date%></td>  
								<td class="right"><%=formatnumber(jiyun_day,0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(add_tax1,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(add_tax2,0)%>&nbsp;</td>
							</tr>
						</tbody>
					</table>
                    <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
						</colgroup>
						<thead>
                            <tr>
							    <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">♣ 주의사항</th>
							</tr>
                            <tr>
							    <th colspan="2" class="first" scope="col">* 주민세 재산분</th>
                                <th colspan="2" scope="col">* 지방소득세 종업원분</th>
							</tr>
                        </thead>
						<tbody>
                            <tr>
							    <td colspan="2" class="left" scope="col">1. 사업장 사용면적(과세대상)이 330㎡이하인 경우는<br>&nbsp;&nbsp;&nbsp;&nbsp;납부하지 않습니다.<br>2. 재산분 주민세는 매년 7월 1일부터 7월 31일까지<br>&nbsp;&nbsp;&nbsp;&nbsp;신고납부합니다.<br>3. 재산분 주민세의 세율은 ㎡당 250원 입니다.</td>  
								<td colspan="2" class="left">1. 상시 고용하는 종업원의 월 평균수가 50인이하인 경우는<br>&nbsp;&nbsp;&nbsp;&nbsp;납부하지 않습니다.<br>2. 종업원분 지방소득세는 급여지급월의 다음달 10일까지<br>&nbsp;&nbsp;&nbsp;&nbsp;신고납부 합니다.<br>3. 종업원분 지방소득세의 세율은 과세대상 급여총액의 0.5%<br>&nbsp;&nbsp;&nbsp;&nbsp;입니다.</td>
							</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	 <td width="20%">
                        <div align=center>
                             <strong class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></strong>
                        </div>
				    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_emp_wetax_printok.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>','insa_pay_emp_wetax_pop','scrollbars=yes,width=1060,height=700')" class="btnType04">출력</a>
					</div>                  
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

