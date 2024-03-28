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

Dim be_pg, curr_date, curr_year, curr_month, curr_day
Dim page, view_condi, ck_sw, pmg_yymm, pmg_emp_name
Dim curr_dd, curr_mm, i, j, title_line
Dim give_date, cal_quarter, cal_month, view_month
Dim cal_year, pgsize, start_page, stpage
Dim rsPay, rsCount, totRecord, total_page
Dim rs_org, from_date, str_param
Dim vOutput, strSql

page = f_Request("page")
view_condi = f_Request("view_condi")
pmg_yymm = f_Request("pmg_yymm")
pmg_emp_name = f_Request("pmg_emp_name")

If view_condi = "" Then
	'view_condi = "케이원"
	curr_dd = CStr(DatePart("d", Now()))
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	pmg_yymm = Mid(CStr(from_date), 1, 4) & Mid(CStr(from_date), 6, 2)
End If

title_line = " 급여지급 현황 "
be_pg = "/pay/insa_pay_mg.asp"

'===================================================
'### Paging
'===================================================
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
str_param = "&view_condi="&view_condi&"&pmg_yymm="&pmg_yymm&"&pmg_emp_name="&pmg_emp_name

'===================================================
'### DB Query & Call Procedure
'===================================================
vOutput = "@totalCnt"	'output return value

'전체 조회 개수
objBuilder.Append "CALL USP_PAY_INSA_PAY_MG_TOTAL('"&pmg_yymm&"' "
objBuilder.Append ", '"&view_condi&"' "
objBuilder.Append ", '"&pmg_emp_name&"' "
objBuilder.Append ", "&vOutput&") "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

strSql = "SELECT "&vOutput
Set rsCount = DBConn.Execute(strSql)

totRecord = CInt(rsCount(vOutput)) 'Result.RecordCount

Call Rs_Close(rsCount)	'RecordSet Close

'리스트 조회
objBuilder.Append "CALL USP_PAY_INSA_PAY_MG_LIST('"&pmg_yymm&"'"
objBuilder.Append ", '"&view_condi&"' "
objBuilder.Append ", '"&pmg_emp_name&"' "
objBuilder.Append ", '"&stpage&"' "
objBuilder.Append ", '"&pgsize&"') "

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'===================================================

curr_date = Mid(CStr(Now()), 1, 10)
curr_year = Mid(CStr(Now()), 1, 4)
curr_month = Mid(CStr(Now()), 6, 2)
curr_day = Mid(CStr(Now()), 9, 2)

' 최근3개년도 테이블로 생성
year_tab(3, 1) = Mid(Now(), 1, 4)
year_tab(3, 2) = CStr(year_tab(3, 1)) & "년"
year_tab(2, 1) = CInt(Mid(Now(), 1, 4)) - 1
year_tab(2, 2) = CStr(year_tab(2, 1)) & "년"
year_tab(1, 1) = CInt(Mid(Now(), 1, 4)) - 2
year_tab(1, 2) = CStr(year_tab(1, 1)) & "년"

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
cal_month = Mid(CStr(Now()), 1, 4) + Mid(CStr(Now()), 6, 2)
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

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.view_condi.value == ""){
					alert ("소속회사을 선택해주세요.");
					return false;
				}
				return true;
			}

			//엑셀다운로드[허정호_20210811]
			function payExcelView(){
				var cnt = $('#totRecord').val();
				console.log(cnt);

				if(parseInt(cnt) === 0){
					alert('검색 내용이 없습니다. 검색 후 다시 시도해 주세요.');
					return false;
				}

				var url = '/pay/excel/insa_excel_pay_pay_report.asp';
				var condi = $('#view_condi').val();
				var yymm = $('#pmg_yymm').val();
				var name = $('#pmg_emp_name').val();
				var param = '?view_condi='+condi+'&pmg_yymm='+yymm+'&pmg_emp_name='+name;

				url += param;

				$(location).attr('href', url);
			}

			//인사기록카드 팝업[허정호_20210811]
			function insaCardPopView(id){
				var url = '/insa/insa_card00.asp';
				var pop_name = '인사 기록 카드';
				var param = '?emp_no='+id;
				var features = 'scrollbars=yes,width=1250,height=670';

				url += param;

				pop_Window(url, pop_name, features);
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
								<label>
									<strong>소속 회사 : </strong>
									<%
									'조직 단위별 SelectBox
									'Call SelectEmpOrgLevel("view_condi", "view_condi", "width:130px", view_condi)
									objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
									objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
									objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

									Set rs_org = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()
									%>
									<select name="view_condi" id="view_condi" type="text" style="width:110px;">
										<option value="">선택</option>
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
									<select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%For i = 24 To 1 Step -1	%>
										<option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i, 1) Then %>selected<%End If  %>><%=month_tab(i, 2)%></option>
                                    <%Next	%>
									</select>
								<strong>성명 : </strong>
								 <input type="text" name="pmg_emp_name" id="pmg_emp_name" value="<%=pmg_emp_name%>" />
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="9%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">최초입사일</th>
                                <th scope="col">입사일</th>
                                <th scope="col">소속</th>
								<th scope="col">기본급</th>
                                <th scope="col">지급액계</th>
                                <th scope="col">공제액계</th>
                                <th scope="col">차인지급액</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim pmg_give_tot, page_cnt, emp_first_date, emp_in_date
						Dim de_deduct_tot, pmg_curr_pay

						If totRecord <= 0 Then
							Response.Write "<tr><td colspan='12' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsPay.EOF
								emp_no = rsPay("pmg_emp_no")
								emp_first_date = rsPay("emp_first_date")
								emp_in_date = rsPay("emp_in_date")
								pmg_give_tot = rsPay("pmg_give_total")
								de_deduct_tot = f_toString(rsPay("de_deduct_total"), 0)
								pmg_curr_pay = pmg_give_tot - de_deduct_tot
	           			%>
							<tr>
								<td class="first"><%=rsPay("pmg_emp_no")%>&nbsp;</td>
                                <td>
									<a href="#" onClick="insaCardPopView('<%=rsPay("pmg_emp_no")%>');"><%=rsPay("pmg_emp_name")%></a>
								</td>
                                <td><%=rsPay("pmg_grade")%>&nbsp;</td>
                                <td><%=rsPay("pmg_position")%>&nbsp;</td>
                                <td><%=emp_first_date%>&nbsp;</td>
                                <td><%=emp_in_date%>&nbsp;</td>
                                <td><%=rsPay("org_name")%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_base_pay"), 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(rsPay("pmg_give_total"), 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(de_deduct_tot, 0)%>&nbsp;</td>
                                <td class="right">
									<a href="#" onClick="pop_Window('/person/insa_pay_person_view.asp?emp_no=<%=rsPay("pmg_emp_no")%>&pmg_yymm=<%=rsPay("pmg_yymm")%>&pmg_company=<%=rsPay("pmg_company")%>','insa_pay_person_pop','scrollbars=yes,width=760,height=700')"><%=FormatNumber(pmg_curr_pay, 0)%></a>&nbsp;
								</td>
                                <td class="left">
								<%
								Call EmpOrgCodeSelect(rsPay("org_code"))
								%>
								</td>
							</tr>
						<%
								rsPay.MoveNext()
							Loop
							Call Rs_Close(rsPay)
						End If
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
					<script>

					</script>
						<div class="btnCenter">
							<a href="#" onclick="payExcelView();" class="btnType04">엑셀다운로드</a>
						</div>
                  	</td>
                  	<td>
					<%
					'Paging 처리[허정호_20210720]
					Call Page_Navi_Ver2(page, be_pg, str_param, totRecord, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				</table>
			</form>
		</div>
	</div>
	<input type="hidden" name="totRecord" id="totRecord" value="<%=totRecord%>" />
	<input type="hidden" name="view_condi" id="view_condi" value="<%=view_condi%>" />
	<input type="hidden" name="pmg_yymm" id="pmg_yymm" value="<%=pmg_yymm%>" />
	<input type="hidden" name="pmg_emp_name" id="pmg_emp_name" value="<%=pmg_emp_name%>" />
	</body>
</html>