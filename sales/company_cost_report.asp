<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'on Error resume next
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
Dim cost_month, sales_saupbu, before_date
Dim condi_sql, mm, cost_year
Dim from_date, end_date, to_date
Dim rsCompCost, arrCompCost
Dim title_line, i, j
Dim view_yn, cost_date

Dim from_month, to_month, min_month, now_month, trade_name

cost_month = f_Request("cost_month")
sales_saupbu = f_Request("sales_saupbu")

from_month = f_Request("from_month")
to_month = f_Request("to_month")

trade_name = f_Request("trade_name")

'If sales_saupbu = "" Then
'	sales_saupbu = "전체"
'End If

'사업부 전체 View 권한
Select Case emp_no
	Case "102592", "100359", "100001", "100740"
		view_yn = "Y"
	Case Else
		view_yn = "N"
		sales_saupbu = bonbu
End Select

'If cost_month = "" Then
'	before_date = DateAdd("m", -1, Now())
'	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date),6,2)
'	sales_saupbu = "전체"
'End If

'min_month = "201501"
now_month = CStr(Mid(Now(), 1, 4)) & CStr(Mid(Now(), 6, 2))

If from_month = "" Then
	from_month = now_month - 1
End If

If to_month = "" Then
	to_month = now_month
End If

'If sales_saupbu = "" Then
'	sales_saupbu = "전체"
'End If

cost_year = Mid(to_month, 1, 4)

'from_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2) & "-01"
'end_date = DateValue(from_date)
'end_date = DateAdd("m", 1, from_date)
'to_date = CStr(DateAdd("d", -1, end_date))
'mm = Mid(cost_month, 5, 2)
'cost_year = Mid(cost_month, 1, 4)
'cost_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2)

title_line = "거래처별 손익현황"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
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
		</script>
		<script type="text/javascript">
			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm() {
				var from_year = $('#from_year').val();
				var to_year = $('#to_year').val();

				if(from_year != to_year){
					alert("검색 년도는 동일해야 합니다.");
					return false;
				}

				if (document.frm.from_month.value == "") {
					alert ("시작년월을 입력하세요.");
					return false;
				}
				if (document.frm.to_month.value == "") {
					alert ("종료년월을 입력하세요.");
					return false;
				}
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}

			//본부 검색
			function saupbuSearch(){
				console.log($('#sales_saupbu').val());

				$('#trade_name').val('');
				frmcheck();
			}
			/*
			function tradeSearch(){
				console.log($('#trade_name').val());

				frmcheck()
			}*/
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3><br/>
				<!--<h3 class="stit">1. 천만원 이하 거래처 비용은 기타 항목으로 처리 </h3>-->
				<form action="/sales/company_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>시작년월&nbsp;</strong>(예201401) :
                                	<input name="from_month" type="text" value="<%=from_month%>" style="width:70px" />
									<input type="hidden" name="from_year" value="<%=Mid(from_month, 1, 4)%>" />
								</label>
								~
								<label>
								&nbsp;&nbsp;<strong>종료년월&nbsp;</strong>(예201501) :
                                	<input name="to_month" type="text" value="<%=to_month%>" style="width:70px" />
									<input type="hidden" name="to_year" value="<%=Mid(to_month, 1, 4)%>" />
								</label>

								<label>
									<strong>사업부 &nbsp;:</strong>
									<%
									Dim rsOrg, arrOrg, org_saupbu

									objBuilder.Append "SELECT saupbu "
									objBuilder.Append "FROM saupbu_sales "
									objBuilder.Append "WHERE saupbu <> '' AND SUBSTRING(sales_date, 1, 4) = '"&cost_year&"' "

									'소속 사업부 조건 처리
									If view_yn = "N" Then
										objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
									End If

									objBuilder.Append "GROUP BY saupbu "
									objBuilder.Append "ORDER BY saupbu ASC "

									Set rsOrg = DBConn.Execute(objBuilder.ToString())

									If Not rsOrg.EOF Then
										arrOrg = rsOrg.getRows()
									End If
									objBuilder.Clear()
									rsOrg.Close() : Set rsOrg = Nothing
									%>
									<select name="sales_saupbu" id="sales_saupbu" style="width:150px" onchange="saupbuSearch();">
										<option value="" <%If sales_saupbu = "" then %>selected<% end if %>>전체</option>
										<%
										If IsArray(arrOrg) Then
											For i = LBound(arrOrg) To UBound(arrOrg, 2)
												org_saupbu = arrOrg(0, i)
										%>
										<option value='<%=org_saupbu%>' <%If org_saupbu = sales_saupbu  then %>selected<% end if %>><%=org_saupbu%></option>
										<%
											Next
										End If
										%>
									</select>
								</label>

								<label>
									<strong>거래처 &nbsp;:</strong>
									<%
									Dim rsTrade, arrTrade, tradeName

									objBuilder.Append "SELECT company_name AS 'trade_name' FROM company_cost_profit "
									objBuilder.Append "WHERE (cost_month >= '"&from_month&"' AND cost_month <= '"&to_month&"') "
									objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
									objBuilder.Append "	AND (sales_cost <> '0' OR (pay_cost + general_cost + common_cost + part_cost + manage_cost)) <> '0' "

									Set rsTrade = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If Not rsTrade.EOF Then
										arrTrade = rsTrade.getRows()
									End If
									rsTrade.Close() : Set rsTrade = Nothing
									%>
									<select name="trade_name" id="trade_name" style="width:150px;" onchange="frmcheck();">
										<option value="" <%If trade_name = "" Then %>selected<%End If %>>전체</option>
										<%
										If IsArray(arrTrade) Then
											For i = LBound(arrTrade) To UBound(arrTrade, 2)
												tradeName = arrTrade(0, i)
										%>
										<option value='<%=tradeName%>' <%If tradeName = trade_name  then %>selected<% end if %>><%=tradeName%></option>
										<%
											Next
										End If
										%>
									</select>
								</label>
								<img src="/image/but_ser.jpg" onclick="frmcheck();" style="cursor:pointer;" alt="검색">
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      			<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="8%" >
							<col width="12%" >
							<col width="12%" >
							<col width="8%" >
							<!--<col width="8%" >-->
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사업부</th>
								<th scope="col">거래처 명</th>
								<th scope="col">매출</th>
								<th scope="col">상주직접비<br/>(인건비)</th>
								<th scope="col">상주직접비<br/>(일반경비)</th>
								<th scope="col">사업부<br/>공통비</th>
								<!--<th scope="col">협업</th>-->
								<th scope="col">부문공통비</th>
								<th scope="col">전사공통비</th>
								<th scope="col">손익</th>
								<th scope="col"></th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="10%" >
							<col width="*%" >
							<col width="8%" >
							<col width="12%" >
							<col width="12%" >
							<col width="8%" >
							<!--<col width="8%" >-->
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="2%" >
						</colgroup>
						<tbody>
						<%
						Dim rsSalesOrg, arrSalesOrg, row_cnt
						Dim company_name, sales_cost, pay_cost, general_cost
						Dim rsComm, comm_cost, direct_cost, rsSalesTot, sales_total
						Dim sales_per, common_cost, common_total, profit_cost
						Dim sales_sum, pay_sum, general_sum, common_sum, profit_sum
						Dim rsSalesCost, company_cost

						Dim rsManage, manage_tot, manage_cost, manage_sum
						Dim rsPart, part_tot_cost, as_tot_cnt, as_cnt, as_saupbu_cnt
						Dim rsSaupbuPart, part_cnt, part_tot
						Dim part_cost, part_sum, comm_per

						Dim rsCompanyTot, company_tot, cowork_cost, cowork_sum
						Dim exceptDate

						'202204월부터 전사공통비 SI1본부 고객사 삼성생명보험(주) 매출 제외 처리(재무 요청)[허정호_20220511]
						exceptDate = "202204"

						sales_sum = 0
						pay_sum = 0
						general_sum = 0
						common_sum = 0
						part_sum = 0
						manage_sum = 0
						profit_sum = 0
						cowork_sum = 0

						'영업 사업부 조회
						objBuilder.Append "SELECT saupbu FROM sales_org "
						objBuilder.Append "WHERE sales_year = '"&cost_year&"' "

						If sales_saupbu <> "" Then
							objBuilder.Append "AND saupbu = '"&sales_saupbu&"' "
						End If

						objBuilder.Append "ORDER BY sort_seq ASC "

						Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsSalesOrg.EOF Then
							arrSalesOrg = rsSalesOrg.getRows()
						End If
						rsSalesOrg.Close() : Set rsSalesOrg = Nothing

						If IsArray(arrSalesOrg) Then
							For i = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
								saupbu = arrSalesOrg(0, i)

								'거래처별 비용 현황
								'objBuilder.Append "CALL USP_SALES_COMPANY_PROFIT_SEL('"&saupbu&"', '"&cost_year&"', '"&MID(from_date, 1, 7)&"', '"&mm&"');"

								objBuilder.Append "SELECT * FROM ("
								objBuilder.Append "SELECT company_name, "
								objBuilder.Append "	SUM(sales_cost) AS 'sales_cost', SUM(pay_cost) AS 'pay_cost', SUM(general_cost) AS 'general_cost', "
								objBuilder.Append "	SUM(common_cost) AS 'common_cost', SUM(part_cost) AS 'part_cost', SUM(manage_cost) AS 'manage_cost', "
								objBuilder.Append "	SUM(profit_cost) AS 'profit_cost', "
								objBuilder.Append "	SUM(pay_cost) + SUM(general_cost) + SUM(common_cost) + SUM(part_cost) + SUM(manage_cost)  AS 'c_cost' "
								'objBuilder.Append "	SUM(cowork_give_cost + cowork_get_cost) AS 'cowork_cost' "
								objBuilder.Append "FROM company_cost_profit "
								objBuilder.Append "WHERE (cost_month >= '"&from_month&"' AND cost_month <= '"&to_month&"') "
								objBuilder.Append "	AND saupbu = '"&saupbu&"' "
								If trade_name <> "" Then
									objBuilder.Append "	AND company_name LIKE '%"&trade_name&"%' "
								End If

								If from_month >= exceptDate Then
									objBuilder.append "AND company_name <> '삼성생명보험(주)' "
								End If

								objBuilder.Append "GROUP BY company_name "
								objBuilder.Append ") r1 WHERE r1.sales_cost <> 0 OR r1.c_cost <> 0 "

								Set rsCompCost = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If Not rsCompCost.EOF Then
									arrCompCost = rsCompCost.getRows()
								End If
								rsCompCost.Close() : Set rsCompCost = Nothing

								If IsArray(arrCompCost) Then
									'리스트 열 개수
									row_cnt = UBound(arrCompCost, 2) + 1

									'사이트 별 분기 처리
									For j = LBound(arrCompCost) To UBound(arrCompCost, 2)
										company_name = arrCompCost(0, j)	'거래처명
										sales_cost = CDbl(f_toString(arrCompCost(1, j), 0))	'거래처별 매출
										pay_cost = CDbl(f_toString(arrCompCost(2, j), 0))	'상주직접비(인건비)
										general_cost = CDbl(f_toString(arrCompCost(3, j), 0))	'상주직접비(일반경비)
										common_cost = CDbl(f_toString(arrCompCost(4, j), 0))	'사업부공통경비
										part_cost = CDbl(f_toString(arrCompCost(5, j), 0))	'부문공통비
										manage_cost = CDbl(f_toString(arrCompCost(6, j), 0))	'사이트별 전사공통비(매출 기준)
										profit_cost = CDbl(f_toString(arrCompCost(7, j), 0))	'NKP 손익
										'cowork_cost = CDbl(f_toString(arrCompCost(9, j), 0))	'협업 비용

										'pay_cost = pay_cost - cowork_cost

										'총계
										sales_sum = FormatNumber(sales_sum + sales_cost, 0)
										pay_sum = FormatNumber(pay_sum + pay_cost, 0)
										general_sum = FormatNumber(general_sum + general_cost, 0)
										common_sum = FormatNumber(common_sum + common_cost, 0)
										part_sum = FormatNumber(part_sum + part_cost, 0)
										manage_sum = FormatNumber(manage_sum + manage_cost, 0)

										profit_sum = FormatNumber(profit_sum + profit_cost, 0)
										'cowork_sum = FormatNumber(cowork_sum + cowork_cost, 0)
							%>
							<!--<tr <%If company_name = "기타" Then %>bgcolor="#FFFFCC"<%End If %>>-->
							<tr>
							<%If j = 0 Then %>
								<td class="first" rowspan="<%=CInt(row_cnt)%>" style="background-color:#EEFFFF;font-weight:bold;"><%=saupbu%></td>
							<%End If %>
								<td><%=company_name%></td>
								<td class="right"><%=FormatNumber(sales_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(pay_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(general_cost, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(common_cost, 0)%>&nbsp;</td>
								<!--<td class="right" style="background-color:#FFFFCC;"><%=FormatNumber(cowork_cost, 0)%>&nbsp;</td>-->
								<td class="right"><%=FormatNumber(part_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(manage_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(profit_cost, 0)%>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
									Next
								End If
							Next
						End If

						DBConn.Close() : Set DBConn = Nothing
						%>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first" style="font-weight:bold;">총계</td>
								<td bgcolor="#FFE8E8" class="right"><%=sales_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=pay_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=general_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=common_sum%>&nbsp;</td>
								<!--<td bgcolor="#FFE8E8" class="right"><%=cowork_sum%>&nbsp;</td>-->
								<td bgcolor="#FFE8E8" class="right"><%=part_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=manage_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=profit_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
							</tr>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
				</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
						<a href="/sales/excel/company_cost_excel.asp?from_month=<%=from_month%>&to_month=<%=to_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">엑셀다운로드</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				</table>
			</form>
			<br>
		</div>
	</div>
	</body>
</html>
