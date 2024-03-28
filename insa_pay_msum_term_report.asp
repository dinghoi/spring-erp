<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_pay_msum_term_report.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	f_yymm=Request.form("from_yymm")
	t_yymm=Request.form("to_yymm")
  else
	view_condi = request("view_condi")
	f_yymm=request("from_yymm")
	t_yymm=Request("to_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	from_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	to_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	f_yymm = from_yymm
	t_yymm = to_yymm
	
	sum_base_pay = 0
	sum_meals_pay = 0
	sum_postage_pay = 0
	sum_re_pay = 0
	sum_overtime_pay = 0
	sum_car_pay = 0
	sum_position_pay = 0
	sum_custom_pay = 0
	sum_job_pay = 0
	sum_job_support = 0
	sum_jisa_pay = 0
	sum_long_pay = 0
	sum_disabled_pay = 0
	sum_family_pay = 0
	sum_school_pay = 0
	sum_qual_pay = 0
	sum_other_pay1 = 0
	sum_other_pay2 = 0
	sum_other_pay3 = 0
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
    sum_nps_amt = 0
    sum_nhis_amt = 0
    sum_epi_amt = 0
    sum_longcare_amt = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_year_incom_tax2 = 0
    sum_year_wetax2 = 0
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	
	
	tax_meals_no = 0	
	tax_car_no = 0	
	tax_meals_yes = 0	
	tax_car_yes = 0	
	
end if

give_date = to_date '지급일

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
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

if view_condi = "전체" then
          Sql = "select * from pay_month_give where (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1')"
	else	  
		  Sql = "select * from pay_month_give where (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1
'if not Rs.eof then
do until rs.eof
	  pay_count = pay_count + 1
							  
	  sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
      sum_meals_pay = sum_meals_pay + int(rs("pmg_meals_pay"))
      sum_postage_pay = sum_postage_pay + int(rs("pmg_postage_pay"))
      sum_re_pay = sum_re_pay + int(rs("pmg_re_pay"))
      sum_overtime_pay = sum_overtime_pay + int(rs("pmg_overtime_pay"))
      sum_car_pay = sum_car_pay + int(rs("pmg_car_pay"))
      sum_position_pay = sum_position_pay + int(rs("pmg_position_pay"))
      sum_custom_pay = sum_custom_pay + int(rs("pmg_custom_pay"))
      sum_job_pay = sum_job_pay + int(rs("pmg_job_pay"))
      sum_job_support = sum_job_support + int(rs("pmg_job_support"))
      sum_jisa_pay = sum_jisa_pay + int(rs("pmg_jisa_pay"))
      sum_long_pay = sum_long_pay + int(rs("pmg_long_pay"))
      sum_disabled_pay = sum_disabled_pay + int(rs("pmg_disabled_pay"))
      sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))
	  
	  'sum_tax_yes = sum_tax_yes + int(rs("pmg_tax_yes"))
	  'sum_tax_no = sum_tax_no + int(rs("pmg_tax_no"))
	  'sum_tax_reduced = sum_tax_reduced + int(rs("pmg_tax_reduced"))
	  
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

if view_condi = "전체" then
          Sql = "select * from pay_month_deduct where (de_yymm >= '"+f_yymm+"' and de_yymm <= '"+t_yymm+"') and (de_id = '1')"
	else	  
		  Sql = "select * from pay_month_deduct where (de_yymm >= '"+f_yymm+"' and de_yymm <= '"+t_yymm+"') and (de_id = '1') and (de_company = '"+view_condi+"')"
end if					  
Set Rs_dct = DbConn.Execute(SQL)							  
							  
'if not Rs_dct.eof then
do until Rs_dct.eof
	  sum_nps_amt = sum_nps_amt + int(Rs_dct("de_nps_amt"))
      sum_nhis_amt = sum_nhis_amt + int(Rs_dct("de_nhis_amt"))
      sum_epi_amt = sum_epi_amt + int(Rs_dct("de_epi_amt"))
      sum_longcare_amt = sum_longcare_amt + int(Rs_dct("de_longcare_amt"))
      sum_income_tax = sum_income_tax + int(Rs_dct("de_income_tax"))
      sum_wetax = sum_wetax + int(Rs_dct("de_wetax"))
	  sum_year_incom_tax = sum_year_incom_tax + int(Rs_dct("de_year_incom_tax"))
      sum_year_wetax = sum_year_wetax + int(Rs_dct("de_year_wetax"))
	  sum_year_incom_tax2 = sum_year_incom_tax2 + int(Rs_dct("de_year_incom_tax2"))
      sum_year_wetax2 = sum_year_wetax2 + int(Rs_dct("de_year_wetax2"))
      sum_other_amt1 = sum_other_amt1 + int(Rs_dct("de_other_amt1"))
      sum_sawo_amt = sum_sawo_amt + int(Rs_dct("de_sawo_amt"))
      sum_hyubjo_amt = sum_hyubjo_amt + int(Rs_dct("de_hyubjo_amt"))
      sum_school_amt = sum_school_amt + int(Rs_dct("de_school_amt"))
      sum_nhis_bla_amt = sum_nhis_bla_amt + int(Rs_dct("de_nhis_bla_amt"))
      sum_long_bla_amt = sum_long_bla_amt + int(Rs_dct("de_long_bla_amt"))	
      sum_deduct_tot = sum_deduct_tot + int(Rs_dct("de_deduct_total"))
	Rs_dct.movenext()
loop
Rs_dct.close()		

sum_curr_pay = sum_give_tot - sum_deduct_tot

curr_yyyy = mid(cstr(f_yymm),1,4)
curr_mm = mid(cstr(f_yymm),5,2)

title_line = cstr(f_yymm) + " ∼ " + cstr(t_yymm) + "월 " + " 급여항목별 집계-" + view_condi 

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
				return "7 1";
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
			<!--#include virtual = "/include/insa_pay_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_msum_term_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where  org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
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
								<strong>귀속년월(시작월) : </strong>
                                    <select name="from_yymm" id="from_yymm" type="text" value="<%=f_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If f_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <label>
								<strong> ∼ 종료월 : </strong>
                                    <select name="to_yymm" id="to_yymm" type="text" value="<%=t_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If t_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="12%" >
                            <col width="12%" >
                            <col width="12%" >
							<col width="*" >
                            <col width="12%" >
                            <col width="12%" >
                            <col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th colspan="4" class="first" style="background:#F5FFFA">지&nbsp;급&nbsp;&nbsp;&nbsp;항&nbsp;목</th>
								<th colspan="4" class="first" style="background:#F8F8FF">공&nbsp;제&nbsp;&nbsp;&nbsp;항&nbsp;목</th>
							</tr>  
                        </thead>
                        <tbody>
							<tr>
								<th class="first" style="background:#F5FFFA">기본급</th>
								<td class="right"><%=formatnumber(sum_base_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">식대(비과세)</th>
								<td class="right"><%=formatnumber(tax_meals_no,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">국민연금</th>
                                <td class="right"><%=formatnumber(sum_nps_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">건강보험</th>
                                <td class="right"><%=formatnumber(sum_nhis_amt,0)%>&nbsp;</td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F5FFFA">통신비</th>
								<td class="right"><%=formatnumber(sum_postage_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">식대</th>
								<td class="right"><%=formatnumber(tax_meals_yes,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">고용보험</th>
                                <td class="right"><%=formatnumber(sum_epi_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">장기요양보험</th>
                                <td class="right"><%=formatnumber(sum_longcare_amt,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#F5FFFA">연장근로수당</th>
								<td class="right"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">소급급여</th>
								<td class="right"><%=formatnumber(sum_re_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">소득세</th>
                                <td class="right"><%=formatnumber(sum_income_tax,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">지방소득세</th>
                                <td class="right"><%=formatnumber(sum_wetax,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#F5FFFA">직책수당</th>
								<td class="right"><%=formatnumber(sum_position_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">주차지원금(비과세)</th>
								<td class="right"><%=formatnumber(tax_car_no,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">기타공제</th>
                                <td class="right"><%=formatnumber(sum_other_amt1,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">경조회비</th>
                                <td class="right"><%=formatnumber(sum_sawo_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">직무보조비</th>
								<td class="right"><%=formatnumber(sum_job_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">주차지원금</th>
								<td class="right"><%=formatnumber(tax_car_yes,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">협조비</th>
                                <td class="right"><%=formatnumber(sum_hyubjo_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">학자금대출</th>
                                <td class="right"><%=formatnumber(sum_school_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">본지사근무비</th>
								<td class="right"><%=formatnumber(sum_jisa_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">고객관리수당</th>
								<td class="right"><%=formatnumber(sum_custom_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">건강보험료정산</th>
                                <td class="right"><%=formatnumber(sum_nhis_bla_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">장기요양보험정산</th>
                                <td class="right"><%=formatnumber(sum_long_bla_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">근속수당</th>
								<td class="right"><%=formatnumber(sum_long_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">업무장려비</th>
								<td class="right"><%=formatnumber(sum_job_support,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">연말정산소득세</th>
                                <td class="right"><%=formatnumber(sum_year_incom_tax,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">연말정산지방세</th>
                                <td class="right"><%=formatnumber(sum_year_wetax,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style=" border-bottom:2px solid #515254; background:#F5FFFA">장애인수당</th>
								<td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_disabled_pay,0)%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td class="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">연말재정산소득세</th>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_year_incom_tax2,0)%>&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">연말재정산지방세</th>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_year_wetax2,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">과세</th>
								<td class="right"><%=formatnumber(sum_tax_yes,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">&nbsp;</th>
								<td class="right">&nbsp;</td>
								<th style="background:#F8F8FF">&nbsp;</th>
                                <td class="left">&nbsp;</td>
                                <th style="background:#F8F8FF">&nbsp;</th>
                                <td class="right">&nbsp;</td>
							</tr>      
                            <tr>
								<th class="first" style="background:#F5FFFA">비과세</th>
								<td class="right"><%=formatnumber(sum_tax_no,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">&nbsp;</th>
								<td class="right">&nbsp;</td>
								<th style="background:#F8F8FF">&nbsp;</th>
                                <td class="left">&nbsp;</td>
                                <th style="background:#F8F8FF">&nbsp;</th>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<th class="first" style="border-bottom:2px solid #515254; background:#F5FFFA">감면소득</th>
								<td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_tax_reduced,0)%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td class="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">&nbsp;</th>
                                <td class="left" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F8F8FF">&nbsp;</th>
                                <td class="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="border-bottom:2px solid #515254; background:#F5FFFA">지급액 계</th>
								<td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_give_tot,0)%>&nbsp;</td>
                                 <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td class="right" style=" border-bottom:2px solid #515254;"><%=pay_count%>&nbsp;명</td>
                                <th style="border-bottom:2px solid #515254; background:#F8F8FF">공제액 계</th>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_deduct_tot,0)%>&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">차인지급액</th>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</td>
							</tr>              
						</tbody>
					</table>
				</div>
                <br>                        
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                    <td width="25%">
					<div class="btnleft">
                    <a href="insa_pay_msum_term_excel.asp?view_condi=<%=view_condi%>&from_yymm=<%=f_yymm%>&to_yymm=<%=t_yymm%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
                    <td width="50%">
                    </td>
				    <td width="25%">
					<div class="btnRight">
                    
					</div>                  
                    </td>        
                    </td> 
			      </tr>
				</table>                
			</form>
		</div>				
	</div>        				
	</body>
</html>

