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
Dim month_tab(24,2)
Dim quarter_tab(8,2)
Dim year_tab(3,2)

Dim be_pg, rever_yyyy
Dim curr_dd, to_date, from_date, curr_mm, i
Dim cal_quarter, cal_month, view_month, j
Dim cal_year, rsPay

be_pg = "/person/insa_plist_pay_mg.asp"
rever_yyyy = f_Request("rever_yyyy")

If rever_yyyy = "" Then
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	rever_yyyy = Mid(CStr(from_date), 1, 4)
End If

' 최근3개년도 테이블로 생성
year_tab(3,1) = Mid(Now(), 1, 4)
year_tab(3,2) = CStr(year_tab(3, 1)) & "년"
year_tab(2,1) = CInt(Mid(Now(), 1, 4)) - 1
year_tab(2,2) = CStr(year_tab(2, 1)) & "년"
year_tab(1,1) = CInt(Mid(Now(), 1, 4)) - 2
year_tab(1,2) = CStr(year_tab(1, 1)) & "년"

' 분기 테이블 생성
curr_mm = Mid(Now(), 6, 2)

If curr_mm > 0 And curr_mm < 4 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) + "1"
End If

If curr_mm > 3 And curr_mm < 7 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) + "2"
End If

If curr_mm > 6 And curr_mm < 10 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) + "3"
End If

If curr_mm > 9 And curr_mm < 13 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) + "4"
End If

quarter_tab(8, 2) = CStr(Mid(quarter_tab(8, 1), 1, 4)) & "년 " & CStr(Mid(quarter_tab(8, 1), 5, 1)) & "/4분기"

For i = 7 To 1 Step -1
	cal_quarter = CInt(quarter_tab(i + 1, 1)) - 1

	If CStr(Mid(cal_quarter, 5, 1)) = "0" Then
		quarter_tab(i, 1) = CStr(CInt(Mid(cal_quarter, 1, 4)) - 1) + "4"
	Else
		quarter_tab(i, 1) = cal_quarter
	End If

	quarter_tab(i, 2) = CStr(Mid(quarter_tab(i, 1), 1, 4)) & "년 " & CStr(Mid(quarter_tab(i, 1), 5, 1)) & "/4분기"
Next

' 년월 테이블생성
cal_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
month_tab(24, 1) = cal_month
view_month = Mid(cal_month, 1, 4) & "년 " & Mid(cal_month, 5, 2) & "월"
month_tab(24, 2) = view_month

For i = 1 To 23
	cal_month = CStr(Int(cal_month) - 1)

	If Mid(cal_month, 5) = "00" Then
		cal_year = CStr(Int(Mid(cal_month, 1, 4)) - 1)
		cal_month = cal_year + "12"
	End If

	view_month = Mid(cal_month, 1, 4) & "년 " & Mid(cal_month, 5, 2) & "월"
	j = 24 - i
	month_tab(j, 1) = cal_month
	month_tab(j, 2) = view_month
Next

Dim title_line, etc_code

title_line = "급여 조회"
etc_code = "9999"

Dim rs_etc, emp_payend_date, emp_payend

objBuilder.Append "SELECT emp_payend_date "
objBuilder.Append "FROM emp_etc_code "
objBuilder.Append "WHERE emp_etc_code = '" & etc_code & "' "

Set rs_etc = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_payend_date = rs_etc("emp_payend_date")
'emp_payend_yn = rs_etc("emp_payend_yn")

rs_etc.Close() : Set rs_etc = Nothing

'사용되는 부분이 없음[허정호_20210720]
'If pmg_yymm > emp_payend_date Then
'	emp_payend = "N"
'Else
'	emp_payend = "Y"
'End If

objBuilder.Append "SELECT pmgt.pmg_yymm, pmgt.pmg_id, pmgt.pmg_emp_no, pmgt.pmg_company, pmgt.pmg_org_name, "
objBuilder.Append "	pmgt.pmg_org_code, pmgt.pmg_grade, pmgt.pmg_base_pay, pmgt.pmg_meals_pay, "
objBuilder.Append "	pmgt.pmg_overtime_pay, pmgt.pmg_give_total, pmgt.pmg_date, pmgt.pmg_position, "
objBuilder.Append "	de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax, de_wetax, "
objBuilder.Append "	de_year_incom_tax, de_year_wetax, de_other_amt1, de_special_tax, de_sawo_amt, "
objBuilder.Append "	de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_total, "
objBuilder.Append "	de_saving_amt, de_johab_amt "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_yymm = de_yymm "
objBuilder.Append "	AND pmdt.de_id = '1' "
objBuilder.Append "	AND pmgt.pmg_emp_no = pmdt.de_emp_no "
objBuilder.Append "WHERE pmgt.pmg_yymm LIKE '%"&rever_yyyy&"%' "
objBuilder.Append "	AND pmgt.pmg_emp_no = '"&emp_no&"' "
objBuilder.Append "	AND pmgt.pmg_id = '1' "

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}

			function delcheck(){
				if(form_chk(document.frm_del)){
					document.frm_del.submit();
				}
			}

			function form_chk(){
				if(!confirm('삭제 하시겠습니까?')) return false;
				else return true;
			}

			//급여 상세 조회[허정호_20210723]
			function payPersView(id){
				console.log(id);
				var url = '/person/insa_pay_person_view.asp';
				var pop_name = '급여 상세 조회';
				var features = 'scrollbars=yes,width=750,height=700';
				var param;

				var arr_str = $('#'+id).val().split('|');
				var emp_no = arr_str[0];
				var yymm = arr_str[1];
				var company = arr_str[2]

				var param = '?emp_no='+emp_no+'&pmg_yymm='+yymm+'&pmg_company='+company;

				url += param;

				pop_Window(url, pop_name, features);
			}
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_plist_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="<%=be_pg%>?ck_sw=n" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                                <label>
								<strong>귀속년도 : </strong>
                                    <select name="rever_yyyy" id="rever_yyyy" type="text" value="<%=rever_yyyy%>" style="width:90px">
                                    <%For i = 3 To 1 Step - 1%>
										<option value="<%=year_tab(i, 1)%>" <%If rever_yyyy = CStr(year_tab(i, 1)) Then %>selected<%End If %>><%=year_tab(i, 2)%></option>
                                    <%Next	%>
                                    </select>
  							    </label>
                                <label>
								<strong>사번 : </strong>
									<input name="emp_no" type="text" value="<%=emp_no%>" style="width:70px" id="emp_no" class="no-input" readonly/>
								</label>
                                <label>
								<strong>성명 : </strong>
									<input name="emp_name" type="text" value="<%=user_name%>" style="width:90px" id="emp_name" class="no-input" readonly/>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">년월</th>
								<th scope="col">소속</th>
								<th scope="col">직급</th>
								<th scope="col">기본급</th>
                                <th scope="col">식대</th>
								<th scope="col">연장수당</th>
                                <th scope="col" style="background:#E0FFFF">지급액계</th>
                                <th scope="col" style="background:#E0FFFF">공제액계</th>
                                <th scope="col" style="background:#FFFFE6">차인지급액</th>
                                <th scope="col">상세조회</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim de_yymm, de_company, de_nps_amt, de_nhis_amt, pmg_give_tot
						Dim de_epi_amt, de_longcare_amt, de_income_tax, de_wetax, de_year_incom_tax
						Dim de_year_wetax, de_other_amt1, de_specil_tx, de_sawo_amt, de_hyubjo_amt
						Dim de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_tot, de_saving_amt
						Dim de_johab_amt, pmg_curr_pay, de_insu_hap, de_tax_hap, de_other_hap, de_bla_amt

						Do Until rsPay.EOF
							de_yymm = rsPay("pmg_yymm")

							If emp_payend_date >= de_yymm Then
								de_company = rsPay("pmg_company")
								pmg_give_tot = rsPay("pmg_give_total")

								de_nps_amt = f_toString(rsPay("de_nps_amt"), 0)
								de_nhis_amt = f_toString(rsPay("de_nhis_amt"), 0)
								de_epi_amt = f_toString(rsPay("de_epi_amt"), 0)
								de_longcare_amt = f_toString(rsPay("de_longcare_amt"), 0)
								de_income_tax = f_toString(rsPay("de_income_tax"), 0)
								de_wetax = f_toString(rsPay("de_wetax"), 0)
								de_year_incom_tax = f_toString(rsPay("de_year_incom_tax"), 0)
								de_year_wetax = f_toString(rsPay("de_year_wetax"), 0)
								de_other_amt1 = f_toString(rsPay("de_other_amt1"), 0)
								de_specil_tx = f_toString(rsPay("de_special_tax"), 0)
								de_sawo_amt = f_toString(rsPay("de_sawo_amt"), 0)
								de_hyubjo_amt = f_toString(rsPay("de_hyubjo_amt"), 0)
								de_school_amt = f_toString(rsPay("de_school_amt"), 0)
								de_nhis_bla_amt = f_toString(rsPay("de_nhis_bla_amt"), 0)
								de_long_bla_amt = f_toString(rsPay("de_long_bla_amt"), 0)
								de_deduct_tot = f_toString(rsPay("de_deduct_total"), 0)

								de_saving_amt = rsPay("de_saving_amt")
								de_johab_amt = rsPay("de_johab_amt")

								pmg_curr_pay = pmg_give_tot - de_deduct_tot
								de_insu_hap = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt
								de_tax_hap = de_income_tax + de_wetax
								de_other_hap = de_other_amt1 + de_sawo_amt + de_hyubjo_amt + de_school_amt
								de_bla_amt = de_nhis_bla_amt + de_long_bla_amt
	           			%>
							<tr>
								<td class="first">
									<%=Mid(CStr(rsPay("pmg_yymm")), 1, 4)%>년&nbsp;<%=Mid(CStr(rsPay("pmg_yymm")), 5, 2)%>월&nbsp;
								</td>
                                <td class="left">
									<%=rsPay("pmg_company")%>&nbsp;-&nbsp;<%=rsPay("pmg_org_name")%>(<%=rsPay("pmg_org_code")%>)&nbsp;
								</td>
                                <td><%=rsPay("pmg_grade")%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_base_pay"), 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_meals_pay"), 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_overtime_pay"), 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_give_total"), 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(de_deduct_tot, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(pmg_curr_pay, 0)%>&nbsp;</td>
                                <td>
									<div>
										<a href="#" onclick="payPersView('<%=emp_no&rsPay("pmg_yymm")%>');">조회</a>
										<input type="hidden" name="<%=emp_no&rsPay("pmg_yymm")%>" id="<%=emp_no&rsPay("pmg_yymm")%>" value="<%=emp_no%>|<%=rsPay("pmg_yymm")%>|<%=rsPay("pmg_company")%>" />
									</div>
								</td>
							</tr>
						<%
							End If

							rsPay.MoveNext()
						Loop
						rsPay.Close() : Set rsPay = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>