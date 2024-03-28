<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_pay_month_ledger.asp"

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
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
    to_date=request("to_date") 
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
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
	
end if

give_date = to_date '지급일

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 분기 테이블 생성
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "년 " + cstr(mid(quarter_tab(8,1),5,1)) + "/4분기"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "년 " + cstr(mid(quarter_tab(i,1),5,1)) + "/4분기"
next

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

rever_yyyymm = mid(cstr(from_date),1,7) '귀속년월
give_date = to_date '지급일

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

'Sql = "SELECT * FROM k1_memb where "+condi_sql+"mg_group = '"+mg_group+"' ORDER BY user_name ASC"
'where_sql = " WHERE isNull(emp_end_date) or emp_end_date = '1900-01-01'"

Sql = "select count(*) from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
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

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then
	    	de_nps_amt = int(Rs_dct("de_nps_amt"))
            de_nhis_amt = int(Rs_dct("de_nhis_amt"))
            de_epi_amt = int(Rs_dct("de_epi_amt"))
		    de_longcare_amt = int(Rs_dct("de_longcare_amt"))
            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
			de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
            de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
            de_other_amt1 = int(Rs_dct("de_other_amt1"))
            de_sawo_amt = int(Rs_dct("de_sawo_amt"))
            de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
            de_school_amt = int(Rs_dct("de_school_amt"))
            de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
            de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
			de_nps_amt = 0
            de_nhis_amt = 0
            de_epi_amt = 0
		    de_longcare_amt = 0
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
			de_year_incom_tax2 = 0
            de_year_wetax2 = 0
            de_other_amt1 = 0
            de_sawo_amt = 0
            de_hyubjo_amt = 0
            de_school_amt = 0
            de_nhis_bla_amt = 0
            de_long_bla_amt = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
     sum_nps_amt = sum_nps_amt + de_nps_amt
     sum_nhis_amt = sum_nhis_amt + de_nhis_amt
     sum_epi_amt = sum_epi_amt + de_epi_amt
	 sum_longcare_amt = sum_longcare_amt + de_longcare_amt
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
	 sum_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
     sum_year_wetax2 = sum_year_wetax2 + de_year_wetax2
     sum_other_amt1 = sum_other_amt1 + de_other_amt1
     sum_sawo_amt = sum_sawo_amt + de_sawo_amt
     sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
     sum_school_amt = sum_school_amt + de_school_amt
     sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
     sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_emp_no ASC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여대장(개인)"

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
				return "0 1";
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
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_month_ledger.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="*" >
							<col width="8%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="7%" >
							<col width="8%" >
                            <col width="7%" >
                            <col width="6%" >
							<col width="6%" > 
                            <col width="6%" >
                            <col width="7%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
				               <th colspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">인적사항</th>
				               <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;">공제 및 차인지급액</th>
			                </tr>
                            <tr>
								<td class="first" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">사번</td> 
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">성  명</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">기본급</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">식대</td>  
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">차량유지비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">통신비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">소급급여</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">연장근로<br>수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">주차지원금</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">국민연금</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">건강보험</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">고용보험</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">장기요양<br>보험료</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">소득세</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">지방소득세</td>
							</tr>
                            <tr>
								<td class="first" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">입사일</td> 
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">직급</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">직책수당</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">고객관리<br>수당</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">직무보조비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">업무장려비</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">본지사<br>근무비</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">근속수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">장애인수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">기타공제</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">사우회<br>회비</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">학자금상환</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">건강보험료<br>정산</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3; font-size:11px">장기요양<br>보험료정산</td>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">공제합계</th>
							</tr>
                            <tr>
								<td class="first" scope="col" style=" border-bottom:2px solid #515254; background:#f8f8f8;">퇴사일</td> 
								<td scope="col" style=" border-bottom:2px solid #515254; background:#f8f8f8;">부서</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<td scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
                                <th scope="col" style=" border-bottom:2px solid #515254;">지급합계</th>
                                <td scope="col" style=" border-bottom:2px solid #515254;">협조비</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말정산<br>소득세</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말정산<br>지방소득세</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말재정산<br>소득세</td>
                                <td scope="col" style=" border-bottom:2px solid #515254;">연말재정산<br>지방세</td>
                                <th scope="col" style=" border-bottom:2px solid #515254; font-size:12px">차인지급액</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  emp_no = rs("pmg_emp_no")
							  pmg_give_tot = rs("pmg_give_total")

	           			%>
							<tr>
								<td class="first"><%=rs("pmg_emp_no")%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_emp_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_meals_pay"),0)%></td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_postage_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_re_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_car_pay"),0)%></td>
                        <%
						      Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not rs_emp.eof then
									emp_first_date = rs_emp("emp_first_date")
									emp_in_date = rs_emp("emp_in_date")
									emp_end_date = rs_emp("emp_end_date")
									emp_company = rs_emp("emp_company")
									emp_bonbu = rs_emp("emp_bonbu")
									emp_saupbu = rs_emp("emp_saupbu")
									emp_team = rs_emp("emp_team")
	                             else
									emp_first_date = ""
									emp_in_date = ""
									emp_company = ""
									emp_bonbu = ""
									emp_saupbu = ""
									emp_team = ""
                              end if
                              rs_emp.close()
							  if emp_end_date = "1999-01-01" then emp_end_date = "" end if
                          %>

                         <%
						      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_nps_amt = int(Rs_dct("de_nps_amt"))
                                    de_nhis_amt = int(Rs_dct("de_nhis_amt"))
                                    de_epi_amt = int(Rs_dct("de_epi_amt"))
		                            de_longcare_amt = int(Rs_dct("de_longcare_amt"))
                                    de_income_tax = int(Rs_dct("de_income_tax"))
                                    de_wetax = int(Rs_dct("de_wetax"))
									de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                                    de_year_wetax = int(Rs_dct("de_year_wetax"))
									de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
                                    de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
                                    de_other_amt1 = int(Rs_dct("de_other_amt1"))
                                    de_sawo_amt = int(Rs_dct("de_sawo_amt"))
                                    de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                                    de_school_amt = int(Rs_dct("de_school_amt"))
                                    de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                                    de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
		                            de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	                             else
									de_nps_amt = 0
                                    de_nhis_amt = 0
                                    de_epi_amt = 0
		                            de_longcare_amt = 0
                                    de_income_tax = 0
                                    de_wetax = 0
									de_year_incom_tax = 0
                                    de_year_wetax = 0
									de_year_incom_tax2 = 0
                                    de_year_wetax2 = 0
                                    de_other_amt1 = 0
                                    de_sawo_amt = 0
                                    de_hyubjo_amt = 0
                                    de_school_amt = 0
                                    de_nhis_bla_amt = 0
                                    de_long_bla_amt = 0
		                            de_deduct_tot = 0
                              end if
                              Rs_dct.close()
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
						  
                          %>
                                <td class="right"><%=formatnumber(de_nps_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_nhis_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_epi_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_longcare_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(de_wetax,0)%></td>
							</tr>
                            <tr>
								<td class="first"><%=emp_in_date%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_grade")%></td>
                                <td class="right"><%=formatnumber(rs("pmg_position_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_custom_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_job_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_job_support"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_jisa_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_long_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_disabled_pay"),0)%></td>
                                <td class="right"><%=formatnumber(de_other_amt1,0)%></td>
                                <td class="right"><%=formatnumber(de_sawo_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_school_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_nhis_bla_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_long_bla_amt,0)%></td>
                                <td class="right"><strong><%=formatnumber(de_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
								<td class="first"><%=emp_end_date%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_org_name")%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><strong><%=formatnumber(rs("pmg_give_total"),0)%></strong></td>
                                <td class="right"><%=formatnumber(de_hyubjo_amt,0)%></td>
                                <td class="right"><%=formatnumber(de_year_incom_tax,0)%></td>
                                <td class="right"><%=formatnumber(de_year_wetax,0)%></td>
                                <td class="right"><%=formatnumber(de_year_incom_tax2,0)%></td>
                                <td class="right"><%=formatnumber(de_year_wetax2,0)%></td>
                                <td class="right"><strong><%=formatnumber(pmg_curr_pay,0)%></strong></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot - sum_deduct_tot
						
						%>
                          	<tr>
                                <td rowspan="3" class="first" style="background:#ffe8e8;">총계</td>
                                <td rowspan="3" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(pay_count,0)%>&nbsp;명</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_base_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_meals_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_postage_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_re_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_overtime_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sumpmg_car_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nps_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_epi_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_longcare_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_income_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_wetax,0)%></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_position_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_custom_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_job_support,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_jisa_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_disabled_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_other_amt1,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_sawo_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_school_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_bla_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_long_bla_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3; font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_give_tot,0)%></strong></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_hyubjo_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_incom_tax2,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_year_wetax2,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=formatnumber(sum_curr_pay,0)%></strong></td>
							</tr>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_pay_month_ledger.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_month_ledger.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_month_ledger.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_month_ledger.asp?page=<%=i%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_month_ledger.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_month_ledger.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_month_ledger_print2.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>','insa_pay_month_ledger_pop','scrollbars=yes,width=1250,height=700')" class="btnType04">급여대장 출력</a>
					</div>                  
                    </td>                    
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

