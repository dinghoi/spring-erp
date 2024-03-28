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

Dim page, view_condi, view_bank, pmg_yymm, to_date, pmg_id, be_pg
Dim curr_dd, from_date, sum_base_pay, sum_meals_pay, sum_postage_pay
Dim sum_give_tot, sum_deduct_tot, sum_curr_pay

Dim give_date, curr_mm, i, cal_quarter
Dim cal_month, view_month, j, cal_year, rever_yyyymm
Dim pgsize, start_page, stpage
Dim rsCount, total_record, pg_url, rsBank

Dim pmg_give_tot, de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
Dim de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2
Dim de_other_amt1, de_sawo_amt, de_hyubjo_amt, de_school_amt, de_nhis_bla_amt
Dim de_long_bla_amt, de_deduct_tot, pmg_give_total, de_deduct_total

Dim rsPay, arrPay, pmg_emp_no, emp_in_date, emp_jikmu, pmg_curr_pay
Dim curr_yyyy, title_line

Dim rs_org, rs_etc

page = f_Request("page")
view_condi = f_Request("view_condi")
view_bank = f_Request("view_bank")
pmg_yymm = f_Request("pmg_yymm")
to_date = f_Request("to_date")
pmg_id = f_Request("pmg_id")

be_pg = "/pay/insa_pay_bank_transfer.asp"

If f_toString(view_condi, "") = "" Then
	view_condi = "케이원"
'	view_bank = "신한은행"
    view_bank = "전체"
	pmg_id = "1"
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	pmg_yymm = Mid(CStr(from_date), 1, 4)&Mid(CStr(from_date), 6, 2)

	sum_give_tot = 0
	sum_deduct_tot = 0
	sum_curr_pay = 0
End If

give_date = to_date '지급일

' 최근3개년도 테이블로 생성
year_tab(3, 1) = Mid(Now(), 1, 4)
year_tab(3, 2) = CStr(year_tab(3, 1))&"년"
year_tab(2, 1) = CInt(Mid(Now(),1, 4)) - 1
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
	cal_quarter = cint(quarter_tab(i+1,1)) - 1

	If CStr(Mid(cal_quarter, 5, 1)) = "0" Then
		quarter_tab(i, 1) = CStr(CInt(Mid(cal_quarter, 1, 4)) - 1)&"4"
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
	cal_month = CStr(CLng(cal_month) - 1)

	If Mid(cal_month, 5) = "00" Then
		cal_year = CStr(CLng(Mid(cal_month, 1, 4)) - 1)
		cal_month = cal_year&"12"
	End If

	view_month = Mid(cal_month, 1, 4)&"년 "&Mid(cal_month, 5, 2)&"월"
	j = 24 - i
	month_tab(j, 1) = cal_month
	month_tab(j, 2) = view_month
Next

rever_yyyymm = Mid(CStr(from_date), 1, 7) '귀속년월
give_date = to_date '지급일

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If
stpage = CLng((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&pmg_yymm="&pmg_yymm&"&pmg_id="&pmg_id&"&view_bank="&view_bank&"&to_date="&to_date

'리스트 조회 카운트
objBuilder.Append "SELECT COUNT(*) FROM pay_month_give "
objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '"&pmg_id&"' "
objBuilder.Append "	AND pmg_company = '"&view_condi&"' "

If view_bank <> "전체" Then
	objBuilder.Append "AND pmg_bank_name = '"&view_bank&"' "
End If

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

curr_yyyy = Mid(CStr(pmg_yymm), 1, 4)
curr_mm = Mid(CStr(pmg_yymm), 5, 2)

title_line = CStr(curr_yyyy)&"년 "&CStr(curr_mm)&"월 "&" 급여 은행별 이체현황"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			/*
		    $(function(){
				$( "#datepicker" ).datepicker();
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
				if (document.frm.view_condi.value == ""){
					alert("소속을 선택하시기 바랍니다");
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
									<select name="view_condi" id="view_condi" type="text" style="width:130px;">
                				<%
								Do Until rs_org.EOF
			  					%>
                						<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
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
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%Next	%>
                                 </select>
								</label>
								<label>

                                <strong>소득구분</strong>
                                <select name="pmg_id" id="pmg_id" type="text" value="<%=pmg_id%>" style="width:100pxl;">
                                    <option value="1" <%If pmg_id = "1" Then %>selected<%End If %>>급여</option>
                                    <option value="2" <%If pmg_id = "2" Then %>selected<%End If %>>상여금</option>
                                    <option value="3" <%If pmg_id = "3" Then %>selected<%End If %>>추천인인센티브</option>
                                    <option value="4" <%If pmg_id = "4" Then %>selected<%End If %>>연차수당</option>
                                </select>
                                </label>
                            <strong>이체은행 : </strong>
								<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '50' ORDER BY emp_etc_name ASC;"

					            Set rs_etc = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
                                <label>
								<select name="view_bank" id="view_bank" type="text" style="width:100px;">
                                    <option value="전체" <%If view_bank = "전체" Then %>selected<%End If %>>전체</option>
                				<%
								Do Until rs_etc.EOF
			  					%>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If view_bank = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                				<%
									rs_etc.MoveNext()
								Loop
								rs_etc.Close() : Set rs_etc = Nothing
								%>
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
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="12%" >
                            <col width="8%" >
                            <col width="10%" >
                            <col width="14%" >
                            <col width="10%" >
							<col width="9%" >
                            <col width="9%" >
						</colgroup>
						<thead>
							<tr>
				               <th class="first" scope="col">사원번호</th>
                               <th scope="col">성명</th>
                               <th scope="col">입사일</th>
                               <th scope="col">직급</th>
                               <th scope="col">소속</th>
                               <th scope="col">직무</th>
				               <th scope="col">이체은행</th>
                               <th scope="col">계좌번호</th>
                               <th scope="col">예금주명</th>
                               <th scope="col">차인지급액</th>
                               <th scope="col">실지급액</th>
			                </tr>
						</thead>
						<tbody>
						<%
						'급여 정보 조회
						objBuilder.Append "SELECT pmgt.pmg_emp_no, pmgt.pmg_give_total, pmgt.pmg_emp_name, pmgt.pmg_grade, "
						objBuilder.Append "	pmgt.pmg_org_name, pmgt.pmg_bank_name, pmgt.pmg_account_no, pmgt.pmg_account_holder, "

						objBuilder.Append "	emtt.emp_in_date, emtt.emp_jikmu, "

						objBuilder.Append "	pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, pmdt.de_income_tax, "
						objBuilder.Append "	pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, pmdt.de_year_wetax2, "
						objBuilder.Append "	pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, pmdt.de_school_amt, pmdt.de_nhis_bla_amt, "
						objBuilder.Append "	pmdt.de_long_bla_amt, pmdt.de_deduct_total, "

						objBuilder.Append "(SELECT SUM(pmg_give_total) FROM pay_month_give "
						objBuilder.Append "WHERE pmg_yymm = pmgt.pmg_yymm AND pmg_id = '1' AND pmg_company = pmgt.pmg_company) AS 'pmg_give_tot', "

						objBuilder.Append "(SELECT SUM(de_deduct_total) FROM pay_month_deduct "
						objBuilder.Append "WHERE de_yymm = pmgt.pmg_yymm AND de_id = '1' AND de_company = pmgt.pmg_company) AS 'de_deduct_tot' "

						objBuilder.Append "FROM pay_month_give AS pmgt "
						objBuilder.Append "INNER JOIN emp_master AS emtt ON pmgt.pmg_emp_no = emtt.emp_no "
						objBuilder.Append "	AND (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' Or emtt.emp_end_date = '') "
						objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
						objBuilder.Append "	AND pmgt.pmg_company = pmdt.de_company "
						objBuilder.Append "	AND pmdt.de_id = '1' AND de_yymm = '"&pmg_yymm&"' "
						objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '"&pmg_id&"' AND pmg_company = '"&view_condi&"' "

						If view_bank <> "전체" Then
							objBuilder.Append "AND pmgt.pmg_bank_name = '"&view_bank&"'"
						End If

						objBuilder.Append "ORDER BY pmgt.pmg_company, pmgt.pmg_bank_name, pmgt.pmg_emp_no ASC "
						objBuilder.Append "LIMIT "&stpage& "," &pgsize&";"

						Set rsPay = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsPay.EOF Then
							Do Until rsPay.EOF
								pmg_emp_no = rsPay("pmg_emp_no")
								pmg_give_total = rsPay("pmg_give_total")

								emp_in_date = f_toString(rsPay("emp_in_date"), "")
								emp_jikmu = f_toString(rsPay("emp_jikmu"), "")

								de_nps_amt = CLng(f_toString(rsPay("de_nps_amt"), 0))
								de_nhis_amt = CLng(f_toString(rsPay("de_nhis_amt"), 0))
								de_epi_amt = CLng(f_toString(rsPay("de_epi_amt"), 0))
								de_longcare_amt = CLng(f_toString(rsPay("de_longcare_amt"), 0))
								de_income_tax = CLng(f_toString(rsPay("de_income_tax"), 0))
								de_wetax = CLng(f_toString(rsPay("de_wetax"), 0))
								de_year_incom_tax = CLng(f_toString(rsPay("de_year_incom_tax"), 0))
								de_year_wetax = CLng(f_toString(rsPay("de_year_wetax"), 0))
								de_year_incom_tax2 = CLng(f_toString(rsPay("de_year_incom_tax2"), 0))
								de_year_wetax2 = CLng(f_toString(rsPay("de_year_wetax2"), 0))
								de_other_amt1 = CLng(f_toString(rsPay("de_other_amt1"), 0))
								de_sawo_amt = CLng(f_toString(rsPay("de_sawo_amt"), 0))
								de_hyubjo_amt = CLng(f_toString(rsPay("de_hyubjo_amt"), 0))
								de_school_amt = CLng(f_toString(rsPay("de_school_amt"), 0))
								de_nhis_bla_amt = CLng(f_toString(rsPay("de_nhis_bla_amt"), 0))
								de_long_bla_amt = CLng(f_toString(rsPay("de_long_bla_amt"), 0))
								de_deduct_total = CLng(f_toString(rsPay("de_deduct_total"), 0))

								pmg_give_tot = CLng(f_toString(rsPay("pmg_give_tot"), 0))
								de_deduct_tot = CLng(f_toString(rsPay("de_deduct_tot"), 0))


								pmg_curr_pay = pmg_give_total - de_deduct_total

								sum_give_tot = sum_give_tot + CLng(pmg_give_tot)
								sum_deduct_tot = sum_deduct_tot + CLng(de_deduct_tot)
							%>
								<tr>
									<td class="first"><%=pmg_emp_no%>&nbsp;</td>
									<td><%=rsPay("pmg_emp_name")%>&nbsp;</td>
									<td><%=emp_in_date%>&nbsp;</td>
									<td><%=rsPay("pmg_grade")%>&nbsp;</td>
									<td><%=rsPay("pmg_org_name")%>&nbsp;</td>
									<td><%=emp_jikmu%>&nbsp;</td>
									<td><%=rsPay("pmg_bank_name")%>&nbsp;</td>
									<td><%=rsPay("pmg_account_no")%>&nbsp;</td>
									<td><%=rsPay("pmg_account_holder")%>&nbsp;</td>
									<td class="right"><%=FormatNumber(pmg_curr_pay, 0)%>&nbsp;</td>
									<td class="right"><%=FormatNumber(pmg_curr_pay, 0)%>&nbsp;</td>
								</tr>
							<%
								rsPay.MoveNext()
							Loop
						Else
							Response.Write "<tr><td colspan='11' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						End If
						rsPay.Close() : Set rsPay = Nothing

						sum_curr_pay = sum_give_tot - sum_deduct_tot
						%>
                          	<tr>
                                <th colspan="9" class="first">총계&nbsp;</th>
                                <th class="right"><%=FormatNumber(sum_curr_pay, 0)%>&nbsp;</th>
                                <th class="right"><%=FormatNumber(sum_curr_pay, 0)%>&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>

				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/pay/insa_excel_pay_bank_transfer.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_id=<%=pmg_id%>&view_bank=<%=view_bank%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
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