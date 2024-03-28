<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim month_tab(24, 2)
Dim quarter_tab(8, 2)
Dim year_tab(3, 2)

Dim page, view_condi, be_pg, pmg_yymm, curr_dd, to_date, from_date
Dim sum_base_pay, sum_meals_pay, sum_postage_pay, sum_re_pay, sum_overtime_pay
Dim sum_car_pay, sum_position_pay, sum_custom_pay, sum_job_pay, sum_job_support
Dim sum_jisa_pay, sum_long_pay, sum_disabled_pay, sum_family_pay, sum_school_pay
Dim sum_qual_pay, sum_other_pay1, sum_other_pay2, sum_other_pay3, sum_tax_yes
Dim sum_tax_no, sum_tax_reduced, sum_give_tot, sum_nps_amt, sum_nhis_amt
Dim sum_epi_amt, sum_longcare_amt, sum_income_tax, sum_wetax, sum_year_incom_tax
Dim sum_year_wetax, sum_year_incom_tax2, sum_year_wetax2, sum_other_amt1, sum_sawo_amt
Dim sum_hyubjo_amt, sum_school_amt, sum_nhis_bla_amt, sum_long_bla_amt, sum_deduct_tot
Dim pay_count, sum_curr_pay, give_date, curr_mm, i, cal_quarter, cal_month, view_month
Dim j, cal_year, rever_yyyymm, pgsize, start_page, stpage, rsCount, total_record

Dim curr_yyyy, title_line, rs_org, rsPay
Dim pg_url

Dim emp_first_date, emp_in_date, emp_end_date, emp_bonbu, emp_saupbu, emp_team
Dim pmg_emp_no, pmg_give_tot, de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
Dim de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2, de_other_amt1
Dim de_sawo_amt, de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt
Dim de_deduct_tot, pmg_curr_pay

page = f_Request("page")
view_condi = f_Request("view_condi")
pmg_yymm = f_Request("pmg_yymm")

be_pg = "/pay/insa_pay_month_ledger.asp"

'curr_date = mid(cstr(now()),1,10)
'curr_year = mid(cstr(now()),1,4)
'curr_month = mid(cstr(now()),6,2)
'curr_day = mid(cstr(now()),9,2)

If f_toString(view_condi, "") = "" Then
	view_condi = "케이원"
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	pmg_yymm = Mid(CStr(from_date), 1, 4)&Mid(CStr(from_date), 6, 2)

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
End If

give_date = to_date '지급일

' 최근3개년도 테이블로 생성
year_tab(3, 1) = Mid(Now(), 1, 4)
year_tab(3, 2) = CStr(year_tab(3, 1))&"년"
year_tab(2, 1) = CInt(Mid(Now(), 1, 4)) - 1
year_tab(2, 2) = CStr(year_tab(2, 1))&"년"
year_tab(1, 1) = CInt(Mid(Now(), 1, 4)) - 2
year_tab(1, 2) = CStr(year_tab(1, 1))&"년"

' 분기 테이블 생성
curr_mm = Mid(Now(), 6, 2)

If curr_mm > 0 And curr_mm < 4 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"1"
End If

If curr_mm > 3 And curr_mm < 7 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"2"
End If

If curr_mm > 6 And curr_mm < 10 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"3"
End If

If curr_mm > 9 And curr_mm < 13 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"4"
End If

quarter_tab(8, 2) = CStr(Mid(quarter_tab(8, 1), 1, 4))&"년 "&CStr(Mid(quarter_tab(8, 1), 5, 1))&"/4분기"

For i = 7 To 1 Step -1
	cal_quarter = CInt(quarter_tab(i+1, 1)) - 1

	If CStr(Mid(cal_quarter, 5, 1)) = "0" Then
		quarter_tab(i, 1) = CStr(CInt(Mid(cal_quarter, 1, 4)) - 1)& "4"
	Else
		quarter_tab(i, 1) = cal_quarter
	End If

	quarter_tab(i, 2) = CStr(Mid(quarter_tab(i, 1), 1, 4))&"년 "&CStr(Mid(quarter_tab(i, 1), 5, 1))&"/4분기"
Next

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = Mid(CStr(Now()), 1, 4)&Mid(CStr(Now()), 6, 2)
month_tab(24, 1) = cal_month
view_month = Mid(cal_month, 1, 4)&"년 "&Mid(cal_month, 5, 2)&"월"
month_tab(24, 2) = view_month

For i = 1 To 23
	cal_month = cstr(int(cal_month) - 1)

	If Mid(cal_month, 5) = "00" Then
		cal_year = CStr(Int(Mid(cal_month, 1, 4)) - 1)
		cal_month = cal_year&"12"
	End If

	view_month = Mid(cal_month, 1, 4)&"년 "&Mid(cal_month, 5, 2)&"월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
Next

rever_yyyymm = Mid(CStr(from_date),1,7) '귀속년월
give_date = to_date '지급일

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = CInt((page-1)*pgsize)
pg_url = "&view_condi="&view_condi&"&pmg_yymm="&pmg_yymm&"&to_date="&to_date

'Sql = "select count(*) from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
objBuilder.Append "SELECT COUNT(*) FROM pay_month_give "
objBuilder.Append "WHERE pmg_yymm='"&pmg_yymm&"' AND pmg_id='1' AND pmg_company='"&view_condi&"';"

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount= Nothing

curr_yyyy = Mid(CStr(pmg_yymm),1,4)
curr_mm = Mid(CStr(pmg_yymm),5,2)
title_line = CStr(curr_yyyy)&"년 "&CStr(curr_mm)&"월 "&" 급여대장(개인)"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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

		    /*$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%'=from_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%'=to_date%>" );
			});*/

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.view_condi.value == ""){
					alert("소속을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								'objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE org_level = '회사' ORDER BY org_code ASC;"
								objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
								objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

								Set rs_org = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px;s">
                			  <%
								Do Until rs_org.EOF
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") Then %>selected<%End If %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.MoveNext()
								Loop
								rs_org.Close() : Set rs_org = Nothing
							  %>
            					</select>
                                </label>
                                <label>
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px;">
                                    <%For i = 24 To 1 Step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) Then %>selected<%End If %>><%=month_tab(i,2)%></option>
                                    <%Next	%>
                                 </select>
								</label>

                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
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
						'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
						objBuilder.Append "SELECT pmgt.pmg_emp_no, pmgt.pmg_give_total, pmgt.pmg_emp_name, pmgt.pmg_grade, "
						objBuilder.Append " pmgt.pmg_org_name, pmgt.pmg_bank_name, pmgt.pmg_account_no, pmgt.pmg_account_holder, "
						objBuilder.Append "	pmgt.pmg_base_pay, pmgt.pmg_meals_pay, pmgt.pmg_postage_pay, pmgt. pmg_re_pay, "
						objBuilder.Append "	pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay, pmgt.pmg_custom_pay, "
						objBuilder.Append "	pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay, "
						objBuilder.Append "	pmgt.pmg_disabled_pay, "

						objBuilder.Append "	emtt.emp_in_date, emtt.emp_jikmu, emtt.emp_first_date, emtt.emp_end_date, emtt.emp_company, "
						objBuilder.Append "	emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "

						objBuilder.Append "	pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, pmdt.de_income_tax, "
						objBuilder.Append "	pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, "
						objBuilder.Append "	pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, "
						objBuilder.Append "	pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, "
						objBuilder.Append "	pmdt.de_school_amt, pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt, pmdt.de_deduct_total, "

						objBuilder.Append "	(SELECT SUM(pmg_give_total) FROM pay_month_give "
						objBuilder.Append "	WHERE pmg_yymm = pmgt.pmg_yymm AND pmg_id = '1' AND pmg_company = pmgt.pmg_company) AS 'pmg_give_tot', "
						objBuilder.Append "	(SELECT SUM(de_deduct_total) FROM pay_month_deduct "
						objBuilder.Append "	WHERE de_yymm = pmgt.pmg_yymm AND de_id = '1' AND de_company = pmgt.pmg_company) AS 'de_deduct_tot' "
						objBuilder.Append "FROM pay_month_give AS pmgt "
						objBuilder.Append "INNER JOIN emp_master AS emtt ON pmgt.pmg_emp_no = emtt.emp_no "
						objBuilder.Append "	AND (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' Or emtt.emp_end_date = '') "
						objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
						objBuilder.Append "	AND pmgt.pmg_company = pmdt.de_company AND pmdt.de_id = '1' AND de_yymm = '"&pmg_yymm&"' "
						objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '1' AND pmg_company = '"&view_condi&"' "
						objBuilder.Append "ORDER BY pmgt.pmg_company, pmgt.pmg_bank_name, pmgt.pmg_emp_no ASC "
						objBuilder.Append "LIMIT "&stpage&","&pgsize&";"

						Set rsPay = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						Do Until rsPay.EOF
							pay_count = pay_count + 1

							emp_first_date = f_toString(rsPay("emp_first_date"), "")
							emp_in_date = f_toString(rsPay("emp_in_date"), "")
							emp_end_date = f_toString(rsPay("emp_end_date"), "")
							emp_company = f_toString(rsPay("emp_company"), "")
							emp_bonbu = f_toString(rsPay("emp_bonbu"), "")
							emp_saupbu = f_toString(rsPay("emp_saupbu"), "")
							emp_team = f_toString(rsPay("emp_team"), "")

							pmg_emp_no = rsPay("pmg_emp_no")
							pmg_give_tot = rsPay("pmg_give_total")

							sum_base_pay = sum_base_pay + CLng(rsPay("pmg_base_pay"))
							sum_meals_pay = sum_meals_pay + CLng(rsPay("pmg_meals_pay"))
							sum_postage_pay = sum_postage_pay + CLng(rsPay("pmg_postage_pay"))
							sum_re_pay = sum_re_pay + CLng(rsPay("pmg_re_pay"))
							sum_overtime_pay = sum_overtime_pay + CLng(rsPay("pmg_overtime_pay"))
							sum_car_pay = sum_car_pay + CLng(rsPay("pmg_car_pay"))
							sum_position_pay = sum_position_pay + CLng(rsPay("pmg_position_pay"))
							sum_custom_pay = sum_custom_pay + CLng(rsPay("pmg_custom_pay"))
							sum_job_pay = sum_job_pay + CLng(rsPay("pmg_job_pay"))
							sum_job_support = sum_job_support + CLng(rsPay("pmg_job_support"))
							sum_jisa_pay = sum_jisa_pay + CLng(rsPay("pmg_jisa_pay"))
							sum_long_pay = sum_long_pay + CLng(rsPay("pmg_long_pay"))
							sum_disabled_pay = sum_disabled_pay + CLng(rsPay("pmg_disabled_pay"))
							sum_give_tot = sum_give_tot + CLng(rsPay("pmg_give_total"))

							de_nps_amt = CLng(rsPay("de_nps_amt"))
							de_nhis_amt = CLng(rsPay("de_nhis_amt"))
							de_epi_amt = CLng(rsPay("de_epi_amt"))
							de_longcare_amt = CLng(rsPay("de_longcare_amt"))
							de_income_tax = CLng(rsPay("de_income_tax"))
							de_wetax = CLng(rsPay("de_wetax"))
							de_year_incom_tax = CLng(rsPay("de_year_incom_tax"))
							de_year_wetax = CLng(rsPay("de_year_wetax"))
							de_year_incom_tax2 = CLng(rsPay("de_year_incom_tax2"))
							de_year_wetax2 = CLng(rsPay("de_year_wetax2"))
							de_other_amt1 = CLng(rsPay("de_other_amt1"))
							de_sawo_amt = CLng(rsPay("de_sawo_amt"))
							de_hyubjo_amt = CLng(rsPay("de_hyubjo_amt"))
							de_school_amt = CLng(rsPay("de_school_amt"))
							de_nhis_bla_amt = CLng(rsPay("de_nhis_bla_amt"))
							de_long_bla_amt = CLng(rsPay("de_long_bla_amt"))
							de_deduct_tot = CLng(rsPay("de_deduct_total"))

							If emp_end_date = "1999-01-01" Then
								esmp_end_date = ""
							End if

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

							pmg_curr_pay = pmg_give_tot - de_deduct_tot
	           			%>
							<tr <%If pay_count Mod 2 = 0 Then %>style="background-color:#EAEAEA;"<%End If%>>
								<td class="first"><%=pmg_emp_no%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rsPay("pmg_emp_name")%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_base_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_meals_pay"),0)%></td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_postage_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_re_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_overtime_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_car_pay"),0)%></td>

                                <td class="right"><%=FormatNumber(de_nps_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_nhis_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_epi_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_longcare_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_income_tax,0)%></td>
                                <td class="right"><%=FormatNumber(de_wetax,0)%></td>
							</tr>
                            <tr <%If pay_count Mod 2 = 0 Then %>style="background-color:#EAEAEA;"<%End If%>>
								<td class="first"><%=emp_in_date%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rsPay("pmg_grade")%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_position_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_custom_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_job_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_job_support"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_jisa_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_long_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_disabled_pay"),0)%></td>
                                <td class="right"><%=FormatNumber(de_other_amt1,0)%></td>
                                <td class="right"><%=FormatNumber(de_sawo_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_school_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_nhis_bla_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_long_bla_amt,0)%></td>
                                <td class="right"><strong><%=FormatNumber(de_deduct_tot,0)%></strong></td>
							</tr>
                            <tr <%If pay_count Mod 2 = 0 Then %>style="background-color:#EAEAEA;"<%End If%>>
								<td class="first"><%=emp_end_date%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rsPay("pmg_org_name")%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><strong><%=FormatNumber(rsPay("pmg_give_total"),0)%></strong></td>
                                <td class="right"><%=FormatNumber(de_hyubjo_amt,0)%></td>
                                <td class="right"><%=FormatNumber(de_year_incom_tax,0)%></td>
                                <td class="right"><%=FormatNumber(de_year_wetax,0)%></td>
                                <td class="right"><%=FormatNumber(de_year_incom_tax2,0)%></td>
                                <td class="right"><%=FormatNumber(de_year_wetax2,0)%></td>
                                <td class="right"><strong><%=FormatNumber(pmg_curr_pay,0)%></strong></td>
							</tr>
						<%
							rsPay.MoveNext()
						Loop
						rsPay.Close() : Set rsPay = Nothing

						sum_curr_pay = sum_give_tot - sum_deduct_tot

						%>
                          	<tr>
                                <td rowspan="3" class="first" style="background:#ffe8e8;">총계</td>
                                <td rowspan="3" class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(pay_count,0)%>&nbsp;명</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_base_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_meals_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_postage_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_re_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_overtime_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_car_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_nps_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_nhis_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_epi_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_longcare_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_income_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_wetax,0)%></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3;font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_position_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_custom_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_job_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_job_support,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_jisa_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_long_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_disabled_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_other_amt1,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_sawo_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_school_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_nhis_bla_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_long_bla_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=FormatNumber(sum_deduct_tot,0)%></strong></td>
							</tr>
                            <tr>
                                <td class="right" style=" border-left:1px solid #e3e3e3; font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=FormatNumber(sum_give_tot,0)%></strong></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_hyubjo_amt,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_year_incom_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_year_wetax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_year_incom_tax2,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=FormatNumber(sum_year_wetax2,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><strong><%=FormatNumber(sum_curr_pay,0)%></strong></td>
							</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/pay/insa_excel_pay_month_ledger.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

