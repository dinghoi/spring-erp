<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim rsComm, rsSales, rs_emp
Dim i, ck_sw, cost_month, before_date, cost_year, cost_mm
Dim title_line

'/include/profit_loss_menu.asp 사용되는 변수 선언
Dim use_id

Dim tot_saupbu_person, tot_saupbu_cost_amt, tot_saupbu_per, tot_saupbu_direct
Dim tot_saupbu_sale, tot_sale_per, tot_saupbu_sale_amt, all_tot_saupbu_sale_amt
Dim tot_cost_amt, tot_charge_per, tot_company_cost, salesDate

Dim prosCost, privCost
Dim costYearMm

ck_sw = Request("ck_sw")

If ck_sw = "y" Then
	cost_month = Request("cost_month")
	saupbu = Request("saupbu")
Else
	cost_month = Request.form("cost_month")
	saupbu = Request.form("saupbu")
End If

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date), 6, 2)
	'costYearMm = Mid(CStr(before_date), 1, 4) & "-" & Mid(CStr(before_date), 6, 2)
'Else
	'costYearMm = Mid(CStr(cost_month), 1, 4) & "-" & Mid(CStr(cost_month), 6, 2)
End If

'해당 년도 별 전망 배부 기준(허정호_20201208)
Select Case Left(cost_month, 4)
	Case "2020"
		prosCost = "0.01179"	'해당 년도 전망 매출
		privCost = "125000"	'해당 년도 월 1인당 비용
	Case "2021"
		prosCost = "0.015696"
		privCost = "168269"
	Case Else	'2019년 까지 사용되는 세팅 값(이전 년도에는 해당값이 없음)
		prosCost = "0.01388"	'해당 년도 전망 매출 / 100만원 기준
		privCost = "133200"	'해당 년도 월 1인당 비용
End Select

cost_year = Mid(cost_month, 1, 4)
cost_mm = Mid(cost_month, 5)

costYearMm = cost_year & "-" & cost_mm

objBuilder.Append "SELECT r.mg_saupbu, /*사업부*/ "
objBuilder.Append "	r.mem_cnt, /*사업부별 전체 인원(급여기준)*/ "
objBuilder.Append "	IFNULL(r1.total_sales, 0) AS total_sales, /*사업부별 총매출*/ "
objBuilder.Append "	r2.saupbu /* 사업부 명 */, "
objBuilder.Append "	IFNULL(r2.saupbu_person, 0) AS saupbu_person /* 사업부 인력(손익제외) */, "
'objBuilder.Append "	IFNULL(r2.tot_person, 0) AS tot_person /* 총인력 */, "
objBuilder.Append "	IFNULL(r2.saupbu_per, 0) AS saupbu_per /* 차지율 */, "
objBuilder.Append "	IFNULL(r2.saupbu_cost_amt, 0) AS saupbu_cost_amt /* 전사공통비1 */, "
objBuilder.Append "	IFNULL(r2.saupbu_sale, 0) AS saupbu_sale /*사업부 매출*/, "
objBuilder.Append "	IFNULL(r2.tot_sale, 0) AS tot_sale /* 총 매출 */, "
objBuilder.Append "	IFNULL(r2.sale_per, 0) AS sale_per /* 차지율 [회사간 거래 제외] */, "
objBuilder.Append "	IFNULL(r2.saupbu_sale_amt, 0) AS saupbu_sale_amt /* 전사공통비2 */, "
objBuilder.Append "	IFNULL(r2.tot_cost_amt, 0) AS tot_cost_amt, "
objBuilder.Append "	IFNULL(r2.all_tot_cost_amt, 0) AS all_tot_cost_amt, "
objBuilder.Append "	IFNULL(r2.direct_cost, 0) AS direct_cost /*직접비*/ "
objBuilder.Append "FROM ( "

objBuilder.Append "	SELECT pmgt.mg_saupbu, COUNT(*) AS mem_cnt "
objBuilder.Append "	FROM pay_month_give AS pmgt "
objBuilder.Append "	INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	WHERE pmgt.pmg_yymm = '"&cost_month&"' "
objBuilder.Append "		AND pmgt.pmg_id = '1' "
objBuilder.Append "		AND emmt.emp_month = '"&cost_month&"' "
objBuilder.Append "		/*AND emmt.cost_except IN ('0','1')*/ "	'손익 구분
objBuilder.Append "	GROUP BY pmgt.mg_saupbu "
objBuilder.Append ") r "

objBuilder.Append "LEFT OUTER JOIN ( "
objBuilder.Append "	SELECT saupbu, IFNULL(SUM(cost_amt), 0) AS total_sales "
objBuilder.Append "	FROM saupbu_sales "
objBuilder.Append "	WHERE SUBSTRING(sales_date, 1, 7) = '"&costYearMm&"' "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 ON r.mg_saupbu = r1.saupbu "

objBuilder.Append "LEFT OUTER JOIN ( "
objBuilder.Append "	SELECT mgct.saupbu /* 사업부 명 */, "
objBuilder.Append "		mgct.saupbu_person /* 사업부 인력 */, "
'objBuilder.Append "		mgct.tot_person /* 총인력 */, "
objBuilder.Append "		mgct.saupbu_per /* 차지율 */, "

objBuilder.Append "		mgct.saupbu_person * "&privCost&" AS saupbu_cost_amt /* 전사공통비1 */, "

objBuilder.Append "		(SELECT IFNULL(SUM(ssa1.cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales AS ssa1 "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(ssa1.sales_date,1,7),'-','') = '"&cost_month&"' "
objBuilder.Append "			AND mgct.saupbu = ssa1.saupbu) AS saupbu_sale, "

objBuilder.Append "		mgct.tot_sale /* 총 매출 */, "

objBuilder.Append "		(SELECT IFNULL(SUM(ssa2.cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales AS ssa2 "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(ssa2.sales_date,1,7),'-','') = '"&cost_month&"' "
objBuilder.Append "			AND mgct.saupbu = ssa2.saupbu) / "
objBuilder.Append "		(SELECT IFNULL(SUM(ssa3.cost_amt), 0) as sales_amt "
objBuilder.Append "		FROM saupbu_sales AS ssa3 "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(ssa3.sales_date, 1, 7), '-', '') = '"&cost_month&"' "
objBuilder.Append "			AND saupbu <> '회사간거래') AS sale_per, "

objBuilder.Append "		(SELECT IFNULL(SUM(ssa4 .cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales AS ssa4 "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(ssa4 .sales_date,1,7),'-','') = '"&cost_month&"' "
objBuilder.Append "			AND mgct.saupbu = ssa4.saupbu) * "&prosCost&" AS saupbu_sale_amt /* 전사공통비2 */, "

objBuilder.Append "		IFNULL(mgct.tot_cost_amt, 0) AS tot_cost_amt, "
objBuilder.Append "		(mgct.saupbu_person * "&privCost&") + ( "
objBuilder.Append "			(SELECT IFNULL(SUM(ssa5.cost_amt), 0) AS sales_amt "
objBuilder.Append "			FROM saupbu_sales AS ssa5 "
objBuilder.Append "			WHERE REPLACE(SUBSTRING(ssa5.sales_date,1,7),'-','') = '"& cost_month &"' "
objBuilder.Append "				AND mgct.saupbu = ssa5.saupbu) "
objBuilder.Append "			* "& prosCost &") AS all_tot_cost_amt, /*전사공통비 합계*/"

objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt_"&cost_mm&"), 0) AS cost_amt "
objBuilder.Append "		FROM company_cost AS sub_cct "
objBuilder.Append "		WHERE cost_year = '"& Left(cost_month, 4) &"' "
objBuilder.Append "			AND sub_cct.cost_center = '직접비' "
objBuilder.Append "			AND sub_cct.saupbu = mgct.saupbu) AS direct_cost /*직접비*/"

objBuilder.Append "	FROM management_cost AS mgct "
objBuilder.Append "	WHERE mgct.cost_month ='"& cost_month &"'"
objBuilder.Append "	GROUP BY mgct.saupbu "
objBuilder.Append "	ORDER BY mgct.saupbu "
objBuilder.Append ") r2 ON r.mg_saupbu = r2.saupbu "

'response.write objBuilder.ToString()


Set rsComm = Server.CreateObject("ADODB.RecordSet")
rsComm.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

If saupbu = "" Then
	If rsComm.EOF Then
		saupbu = ""
	Else
		saupbu = rsComm("saupbu")
	End If
End If

title_line = "공통비 인원 및 매출 배분 기준 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<script src="/java/jquery-1.9.1.js"></script>
		<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			/*
			function getPageCode(){
				return "2 1";
			}*/

			//검색 결과
			function frmcheck(){
				var frm = document.frm;

				if (frm.cost_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}

				//발생년월 유효 검사[허정호_20201209]
				var costMonth = $('#cost_month').val();
				var monthStr = costMonth.substring(4, 6);
				var monthLen = monthStr.length;

				if(monthLen < 2 || monthLen > 2){
					alert("정확한 발생년월을 입력해 주세요.");
					return false;
				}

				if(monthStr > 12 || monthStr < 1){
					alert("정확한 발생년월을 입력해 주세요.");
					return false;
				}

				frm.submit();
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}

				//cost_month 월 유효검사
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<h3 class="stit">1. 전사공통비 배부 기준은 사업부별 손익에는 인원수, 고객사별손익은 해당 사업부내의 매출액 비율로 배부함. </h3>
				<h3 class="stit">2. 고객사별손익에 직접비 배분은 사업부내의 매출액 비율로 배부함. </h3>
				<form action="/sales/management_cost_report.asp" method="post" name="frm">
					<fieldset class="srch">
						<legend>조회영역</legend>
						<dl>
							<dt>조건 검색</dt>
							<dd>
								<p>
									<label>
										&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) :
                                        <input name="cost_month" id="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
									</label>
									<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>

								</p>
							</dd>
						</dl>
					</fieldset>
				</form>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="52%" height="356" valign="top">
				      	<h3 class="stit">* 사업부별 인원 현황 및 비율</h3>
				      	<table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                                <col width="*" >
								<col width="7%" >
                                <col width="7%" >
                                <col width="10%" >
                                <col width="12%" >
                                <col width="12%" >

                                <col width="14%" >
                                <col width="10%" >
                                <col width="12%" >
                            </colgroup>
				        	<thead>
                            <tr>
                                <th class="first" scope="col" rowspan="2">사업부</th>
								<th class="right" scope="col" rowspan="2" style="text-align:center;">전사 인원<br/>(급여 기준)</th>
                                <th scope="col" colspan="4" style="border-bottom:1px solid #e3e3e3;">전사공통비(인원)</th>
                                <th scope="col" colspan="3" style="border-bottom:1px solid #e3e3e3;">전사공통비(매출)</th>
                                <th scope="col" rowspan="2" style="border-bottom:1px solid #e3e3e3;">전사공통비합계</th>
                            </tr>
                            <tr>
                                <th scope="col" style="border-left:1px solid #e3e3e3;">사업부 인력<br/>(손익 포함)</th>
                                <th scope="col">차지율(%)</th>
                                <th scope="col">전사공통비</th>
                                <th scope="col">직접비</th>

                                <th scope="col">사업부매출</th>
                                <th scope="col">차지율(%)</th>
                                <th scope="col">전사공통비</th>
                            </tr>
                            </thead>
			                <tbody>
			            	<%
							Dim tot_emp_person
							tot_emp_person = 0
                            tot_saupbu_person   = 0
                            tot_saupbu_cost_amt = 0
                            tot_saupbu_per      = 0
                            tot_saupbu_direct   = 0

                            tot_saupbu_sale     = 0
                            tot_sale_per        = 0
                            tot_saupbu_sale_amt = 0
                            all_tot_saupbu_sale_amt = 0

                            i = 0
                            Do Until rsComm.EOF
                                i = i + 1

                                'saupbu_tab(i,1) = rs("saupbu")
                                'saupbu_tab(i,2) = CDbl(rs("direct_cost"))

								tot_emp_person = tot_emp_person + CDbl(rsComm("mem_cnt"))

                                tot_saupbu_person   = tot_saupbu_person + CDbl(rsComm("saupbu_person"))
                                tot_saupbu_cost_amt = tot_saupbu_cost_amt + CDbl(rsComm("saupbu_cost_amt"))
                                tot_saupbu_per      = tot_saupbu_per + rsComm("saupbu_per")
								tot_saupbu_direct   = tot_saupbu_direct + CDbl(rsComm("direct_cost"))

                                tot_saupbu_sale     = tot_saupbu_sale + rsComm("saupbu_sale")
                                tot_sale_per        = tot_sale_per + rsComm("sale_per")
                                tot_saupbu_sale_amt = tot_saupbu_sale_amt + rsComm("saupbu_sale_amt")
								all_tot_saupbu_sale_amt = all_tot_saupbu_sale_amt+ rsComm("all_tot_cost_amt")
                                %>
                                <tr>
                                    <!--사업부     -->
									<td class="first">
									<%
									If rsComm("mg_saupbu") = "" Then
										Response.Write "기타"
									Else
										Response.Write rsComm("mg_saupbu")
									End If
									%>
									</td>
									<td class="right">
										<a href="#" onclick="pop_Window('./pop_dept_person.asp?dept=<%=rsComm("mg_saupbu")%>&dt=<%=cost_month%>','전사 인력 리스트','scrollbars=yes,width=800px,height=700px')"><%=FormatNumber(rsComm("mem_cnt"), 0)%></a>&nbsp;
									</td>
									<!--전사공통비(인원)-->
                                    <!--사업부인력 -->
									<td class="right"><%'=rsComm("saupbu_person")%>
										<a href="#" onclick="pop_Window('./pop_dept_person_comm.asp?dept=<%=rsComm("mg_saupbu")%>&dt=<%=cost_month%>','인력 리스트','scrollbars=yes,width=500px,height=700px')"><%=FormatNumber(rsComm("saupbu_person"), 0)%></a>&nbsp;
									</td>
                                    <!--차지율     -->
									<td class="right"><%=FormatNumber(rsComm("saupbu_per") * 100, 3)%>%&nbsp;</td>
                                    <!--전사공통비 -->
									<td class="right"><%=FormatNumber(rsComm("saupbu_cost_amt"), 0)%>&nbsp;</td>
                                    <!--직접비     -->
									<td class="right"><%=FormatNumber(CDbl(rsComm("direct_cost")), 0)%>&nbsp;</td>

									<!--전사공통비(매출)-->
                                    <!--사업부매출 -->
									<td class="right">
										<a href="#" onclick="pop_Window('./pop_dept_cost.asp?dept=<%=rsComm("mg_saupbu")%>&dt=<%=cost_month%>','사업부내 회사별 매출액 비율','scrollbars=yes,width=800px,height=700px')">
										<%=FormatNumber(rsComm("saupbu_sale"), 0)%>
										</a>
									&nbsp;
									</td>
                                    <!--차지율     -->
									<td class="right"><%=FormatNumber(rsComm("sale_per") * 100, 3)%>%&nbsp;</td>
                                    <!--전사공통비 -->
									<td class="right"><%=FormatNumber(rsComm("saupbu_sale_amt"), 0)%>&nbsp;</td>
                                    <!--전사공통비 합계 -->
									<td class="right"><%=FormatNumber(rsComm("all_tot_cost_amt"), 0)%>&nbsp;</td>
                                </tr>
                                <%
				        	    rsComm.MoveNext()
				        	Loop

				        	rsComm.Close()
							Set rsComm = Nothing
				        	%>
				            <tr bgcolor="#FFE8E8">
								<td class="first">계</td>
								<!--전사인원 총계 -->
								<td class="right"><%=FormatNumber(tot_emp_person, 0)%>&nbsp;</td>
                                <!--사업부인력 계 -->
								<td class="right"><%=FormatNumber(tot_saupbu_person, 0)%>&nbsp;</td>
                                <!--차지율 계 -->
								<td class="right"><%=FormatNumber(tot_saupbu_per * 100, 3)%>%&nbsp;</td>
                                <!--전사공통비 계 -->
								<td class="right"><%=FormatNumber(tot_saupbu_cost_amt, 0)%>&nbsp;</td>
                                <!--직접비 계 -->
								<td class="right"><%=FormatNumber(tot_saupbu_direct, 0)%>&nbsp;</td>

                                <!--사업부매출 계 -->
								<td class="right"><%=FormatNumber(tot_saupbu_sale, 0)%>&nbsp;</td>
                                <!--차지율 계 -->
								<td class="right"><%=FormatNumber(tot_sale_per * 100, 3)%>%&nbsp;</td>
                                <!--전사공통비 계 -->
								<td class="right"><%=FormatNumber(tot_saupbu_sale_amt, 0)%>&nbsp;</td>
                                <!--전사공통비 계 -->
								<td class="right"><%=FormatNumber(all_tot_saupbu_sale_amt, 0)%>&nbsp;</td>
                            </tr>
                            </tbody>
			          </table>
                      </td>

			        </tr>

			      </table>
                </div>
			</div>
	</div>
	</body>
</html>

