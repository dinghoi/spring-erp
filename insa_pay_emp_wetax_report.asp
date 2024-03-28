<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_pay_emp_wetax_report.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	'pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	
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
	
end if

' 년월 테이블생성
cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

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

pmg_date = curr_date '테스트

sum_give_tot = sum_tax_yes + sum_tax_no

month_person_pay = int(sum_tax_yes / pay_count) '신고월 월적용급여액
deduct_14 = month_person_pay * (pay_count - pay_count) '공제액
income_pay15 = sum_tax_yes - deduct_14 '산출과표
income_tax16 = int(income_pay15 * (0.5 / 100)) '산출세액
add_tax1 = 0
add_tax2 = 0
add_tax17 = 0
tax_hap = income_tax16 + add_tax17

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = " 종업원할사업소세(지방세) "

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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_tax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_emp_wetax_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <h3 class="stit">*종업원분 주민세&nbsp;&nbsp;</h3>
				<div class="gView">
                    <table width="175%" border="0" cellpadding="0" cellspacing="0">
				        <tr>
                            <td width="50%" class="left">&nbsp;&nbsp;&nbsp;&nbsp;귀속년월:&nbsp;<%=mid(pmg_yymm,1,4)%>년&nbsp;<%=mid(pmg_yymm,5,2)%>월분</td>
                            <td width="50%" class="right">급여지급일:&nbsp;<%=pmg_date%></td>
                        </tr>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
						</colgroup>
						<thead>
							<tr>
				                <th rowspan="2" class="first" scope="col">구분</th>
                                <th rowspan="2" scope="col">8.사업소인원</th>
				                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">과세표준액</th>
			                </tr>
                            <tr>
							    <th scope="col" style=" border-left:1px solid #e3e3e3;">10.과세제외급여</th>
								<th scope="col">11.과세급여</th>  
								<th scope="col">9.총지급급여액</th>
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
                <h3 class="stit">중소기업 고용지원에 해당되는 중소기업의 공제액(지방세법 제84조의5에 해당하는경우)</h3>    
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="40%" >
                            <col width="30%" >
                            <col width="30%" >
						</colgroup>
						<thead>
                            <tr>
							    <th class="first" scope="col">12.직전연도 월평균 종업원수</th>
								<th scope="col">13.신고월 월적용급여액(11/8)</th>  
								<th scope="col">14.공제액(13*(8-12))</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                                <td class="right"><%=formatnumber(pay_count,0)%>&nbsp;원&nbsp;</td>
                                <td class="right"><%=formatnumber(month_person_pay,0)%>&nbsp;원&nbsp;</td>
                                <td class="right"><%=formatnumber(deduct_14,0)%>&nbsp;원&nbsp;</td>
							</tr>
						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
						</colgroup>
						<thead>
                            <tr>
							    <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">15.산출과표(11-14)</th>  
								<td class="right"><%=formatnumber(income_pay15,0)%>&nbsp;원&nbsp;</td>
								<th scope="col" style=" border-bottom:1px solid #e3e3e3;">16.산출세액(15*0.5%)</th> 
                                <td class="right"><%=formatnumber(income_tax16,0)%>&nbsp;원&nbsp;</td>
							</tr>
                            <tr>
							    <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">납부불성실가산세</th>  
								<td class="right"><%=formatnumber(add_tax1,0)%>&nbsp;원&nbsp;</td>
								<th scope="col" style=" border-bottom:1px solid #e3e3e3;">신고불성실가산세</th> 
                                <td class="right"><%=formatnumber(add_tax2,0)%>&nbsp;원&nbsp;</td>
							</tr>
                            <tr>
							    <th class="first" scope="col">17.가산세</th>  
								<td class="right"><%=formatnumber(add_tax17,0)%>&nbsp;원&nbsp;</td>
								<th scope="col">신고세액합계(16+17)</th> 
                                <td class="right"><%=formatnumber(tax_hap,0)%>&nbsp;원&nbsp;</td>
							</tr>
						</thead>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_pay_empwetax_report.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_emp_wetax_print.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>','insa_pay_emp_wetax_pop','scrollbars=yes,width=1250,height=600')" class="btnType04">납부서</a>
					</div>                  
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

