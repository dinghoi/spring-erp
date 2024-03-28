<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
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
Dim cost_year, base_year, be_year
Dim view_sw, i, j, k

Dim year_tab(5)
Dim sum_amt(20,3,13)
Dim saupbu_tab(20)

Dim rsSalesDept, rsCostStats, rsSaleStats, rsProfitStats, rsEtcStats
Dim title_line, use_comment
Dim cost_saupbu

Dim arrSalesDept

cost_year = f_Request("cost_year")	'조회 년도

title_line = "사업부별 손익 총괄 현황(표준)"

If cost_year = "" Then
	cost_year = Mid(CStr(Now()),1 , 4)
	base_year = cost_year
	view_sw = "0"
End If

be_year = Int(cost_year) - 1

'검색 조회 년도
For i = 1 To 5
	year_tab(i) = Int(Mid(CStr(Now()), 1, 4)) - i + 1
Next

'For i = 0 To 4
'	year_tab(i) = Int(Mid(CStr(Now()), 1, 4)) + i
'Next

For i = 1 To 20
	saupbu_tab(i) = ""
Next

For i = 1 To 20
	For j = 1 To 3
		For k = 1 To 13
			sum_amt(i, j, k) = 0
		Next
	Next
Next

' 2019.02.22 박정신 요청 '사업부별 손익총괄'에서 해당년도에 사업부를 셋팅하면됨
' 영업조직 발췌
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' AND sort_seq <> '31' "	'OA수행본부는 제외

If team="회계재무" Or user_id = "102592" Then
	objBuilder.Append "ORDER BY sort_seq ASC "	' 회계재무 일때문 기타사업부가 들어가도록 하자..
Else
	objBuilder.Append "	AND saupbu <> '기타사업부' "
	objBuilder.Append "ORDER BY sort_seq ASC "
End If

Set rsSalesDept = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'i = 0

'Do Until rsSalesDept.EOF
'	i = i + 1
'	saupbu_tab(i) = rsSalesDept("saupbu")

'	rsSalesDept.MoveNext()
'Loop

If Not rsSalesDept.EOF Then
	arrSalesDept = rsSalesDept.getRows()
End If
rsSalesDept.Close() : Set rsSalesDept = Nothing

If IsArray(arrSalesDept) Then
	For i = LBound(arrSalesDept) To UBound(arrSalesDept, 2)
		saupbu_tab(i + 1) = arrSalesDept(0, i)
	Next
End If

'---------------------------------------------------------------------------------------------------------------
'// 2017-09-15 회계재무 팀만 기타사업부,회사간거래 조회 가능하게 수정
'---------------------------------------------------------------------------------------------------------------

If team="회계재무" Or user_id = "102592" Then
	'i = i + 1
	'saupbu_tab(i) = "기타사업부"
	'i = i + 1
	'saupbu_tab(i) = "회사간거래"
	'i = i + 1
'	saupbu_tab(i) = "솔루션사업부"

	' 회사간거래
	'sql = "select cost_center,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where cost_year = '"&cost_year&"' and (cost_center = '회사간거래') group by cost_center"
	objBuilder.Append "SELECT cost_center, SUM(cost_amt_01), SUM(cost_amt_02), "
	objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "
	objBuilder.Append "	SUM(cost_amt_06), SUM(cost_amt_07), SUM(cost_amt_08), "
	objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "
	objBuilder.Append "	SUM(cost_amt_12) "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
	objBuilder.Append "	AND cost_center = '회사간거래' "
	objBuilder.Append "GROUP BY cost_center "

	Set rsCostStats = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rsCostStats.EOF
		For k = 1 To 12
			sum_amt(i, 2, k) = sum_amt(i, 2, k) + CDbl(rsCostStats(k))
		Next

		rsCostStats.MoveNext()
	Loop

	rsCostStats.close() : Set rsCostStats = Nothing
End If

'---------------------------------------------------------------------------------------------------------------

' 매출 집계
'sql = "select substring(sales_date,1,7) as sales_month,saupbu,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&cost_year&"' group by substring(sales_date,1,7), saupbu"

objBuilder.Append "SELECT SUBSTRING(sales_date, 1, 7) AS sales_month, "
objBuilder.Append "	saupbu,	SUM(cost_amt) AS cost  "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date,1,4) = '"&cost_year&"' "
objBuilder.Append "GROUP BY SUBSTRING(sales_date,1,7), saupbu "

'objBuilder.Append "SELECT SUBSTRING(sast.sales_date, 1, 7) AS sales_month, "
'objBuilder.Append "	eomt.org_bonbu AS saupbu, "
'objBuilder.Append "	SUM(sast.cost_amt) AS cost "
'objBuilder.Append "FROM saupbu_sales AS sast "
'objBuilder.Append "INNER JOIN emp_master AS emtt ON sast.emp_no = emtt.emp_no "
'objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
'objBuilder.Append "WHERE SUBSTRING(sast.sales_date, 1, 4) = '"&cost_year&"' "
'objBuilder.Append "GROUP BY SUBSTRING(sales_date,1,7), eomt.org_bonbu "

Set rsSaleStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsSaleStats.EOF
	For i = 1 To 20
		If saupbu_tab(i) = rsSaleStats("saupbu") Then
			j = 1
			k = Int(Mid(rsSaleStats("sales_month"), 6, 2))

			sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsSaleStats("cost"))

			Exit For
		End If
	Next

	rsSaleStats.MoveNext()
Loop

rsSaleStats.Close() : Set rsSaleStats = Nothing

Dim arrManage, arrManageCost, arrComm, arrCommCost
Dim kk, manage_cost, comm_cost, manage_total, comm_total

'부문 공통비 배부 기준 및 예상 비용
arrManage = Array("SI1본부", "SI2본부", "NI본부", "공공본부")
arrManageCost = Array("115500000", "50200000", "35300000", "400000")

'전사 공통비 배부 기준 및 예상 비용
arrComm = Array("SI1본부", "SI2본부", "NI본부", "공공본부", "ICT본부", "금융SI본부", "공공SI본부", "스마트본부", "DI사업부문")
arrCommCost = Array("78000000", "83000000", "30000000", "22000000", "19000000", "20000000", "17000000", "5000000", "5000000")

' 비용 집계
'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' group by saupbu"

objBuilder.Append "SELECT saupbu, SUM(cost_amt_01), SUM(cost_amt_02), "
objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "
objBuilder.Append "	SUM(cost_amt_06), SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "
objBuilder.Append "	SUM(cost_amt_12) "
objBuilder.Append "FROM saupbu_profit_loss "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "	AND cost_center NOT IN ('전사공통비', '부문공통비') "
objBuilder.Append "	AND saupbu IN (SELECT saupbu FROM sales_org WHERE sales_year = '"&cost_year&"' AND sort_seq <> '9') "

'표준 손익에서는 설치공사 계정 제외 -> 협업 설치공사 포함 처리[20220114_허정호]
'objBuilder.Append "	AND cost_detail NOT IN ('설치공사') "
'objBuilder.Append "	AND cost_detail NOT IN ('설치공사', '협업') "

'objBuilder.Append "	AND cost_amt_01 <> 0 "
objBuilder.Append "GROUP BY saupbu "

Set rsProfitStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsProfitStats.EOF
	For i = 1 To 20

		'부문
		manage_cost = 0
		If i < 5 Then
			If saupbu_tab(i) = arrManage(i-1) Then
				manage_cost = arrManageCost(i-1)
			End If
		End If

		'공통
		comm_cost = 0
		If i < 10  Then
			If saupbu_tab(i) = arrComm(i-1) Then
				comm_cost = arrCommCost(i-1)
			End If
		End If

		If saupbu_tab(i) = rsProfitStats("saupbu") Then
			j = 2

			For k = 1 To 12
				'sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsProfitStats(k))
				If CDbl(rsProfitStats(k)) = 0 Then
					sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsProfitStats(k))
				Else
					sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsProfitStats(k)) + manage_cost + comm_cost
				End If

				'Response.write sum_amt(i, j, k) & " | " & CDbl(rsProfitStats(k)) & " | " & manage_cost & " | " & comm_cost & "<br>"
			Next

			Exit For
		End If
	Next

	rsProfitStats.MoveNext()
Loop

rsProfitStats.Close() : Set rsProfitStats = Nothing

' 비용 집계 (기타사업부)
'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and (saupbu = '' or saupbu = '기타사업부') group by saupbu"

objBuilder.Append "SELECT saupbu, SUM(cost_amt_01), SUM(cost_amt_02), "
objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "
objBuilder.Append "	SUM(cost_amt_06), SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "
objBuilder.Append "	SUM(cost_amt_12) "
objBuilder.Append "FROM saupbu_profit_loss "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "	AND (saupbu = '' OR saupbu = '기타사업부') "

objBuilder.Append "	AND cost_center NOT IN ('전사공통비', '부문공통비') "
'objBuilder.Append "	AND cost_amt_01 <> 0 "

objBuilder.Append "GROUP BY saupbu "

Set rsEtcStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsEtcStats.EOF
	cost_saupbu = Trim(rsEtcStats("saupbu")&"")

	If cost_saupbu = "" Then
		cost_saupbu = "기타사업부"
	End If

	For i = 1 To 20
		If saupbu_tab(i) = cost_saupbu Then
			j = 2

			For k = 1 To 12
				sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsEtcStats(k))
			Next

			Exit For
		End If
	Next

	rsEtcStats.MoveNext()
Loop

rsEtcStats.Close() : Set rsEtcStats = Nothing


' 비용이 없으면 매출도 표기 하지 않음
'for i = 1 to 20
'	if saupbu_tab(i) = "" then
'		exit for
'	end if
'	for k = 1 to 12
'		if sum_amt(i,2,k) = 0 then
'			sum_amt(i,1,k) = 0
'		end if
'	next
'next

' 손익계산
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	j = 3
	For k = 1 To 12
		sum_amt(i, j, k) = sum_amt(i, 1, k) - sum_amt(i, 2, k)
	Next
Next

' 년 합계
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	For j = 1 To 3
		For k = 1 To  12
			sum_amt(i, j, 13) = sum_amt(i, j, 13) + sum_amt(i, j, k)
		Next
	Next
Next

' 총계 : sum_amt(본부(0:총계), 구분, 년도)
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	For j = 1 To 3
		For k = 1 To 13
			sum_amt(0,j,k) = sum_amt(0,j,k) + sum_amt(i,j,k)
		Next
	Next
Next
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
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
				if (formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if (document.frm.cost_year.value == ""){
					alert ("조회년도를 입력하세요.");
					return false;
				}
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
				<form action="/sales/saupbu_profit_loss_total_std.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
						<dd>
							<p>
								<label>
									&nbsp;&nbsp;<strong>조회년도&nbsp;</strong> :
									<select name="cost_year" id="cost_year" style="width:70px">
									<%For i = 1 To 5
									'For i = 0 To 4
									%>
										<option value="<%=year_tab(i)%>" <%If CInt(cost_year) = CInt(year_tab(i)) Then%>selected <%End If %>>&nbsp;<%=year_tab(i)%></option>
									<%
									Next
									%>
									</select>
								</label>
								<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
							</p>
						</dd>
					</dl>
				</fieldset>
				<div  style="text-align:right"><strong>금액단위 : 천원</strong></div>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
							  <th class="first" scope="col">본부</th>
							  <th scope="col">구분</th>
							  <%For i = 1 To 12	%>
							  <th scope="col"><%=i%>월</th>
							  <%Next%>
							  <th scope="col">합계</th>
							</tr>
						</thead>
						<tbody>
							<%
							For i = 1 To 20
								If saupbu_tab(i) = "" Then
									Exit For
								End If
							%>
							<tr>
								<td rowspan="3" class="first"><%=saupbu_tab(i)%></td>
								<td>매출</td>
								<%
								For k = 1 To 13
								%>
								<td class="right"><%=FormatNumber(sum_amt(i, 1, k)/1000, 0)%></td>
								<%
								Next
								%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">비용</td>
								<%
								For k = 1 To 13
								%>
								<td class="right">
								<%If(k < 13 And sum_amt(i, 2, k) > 0) And (saupbu_tab(i) <> "회사간거래") Then
									'보안 사항으로 소속 부서 제한 열람 조건 추가[허정혼_20220106]
									If empProfitViewAll = "Y" Then	'전체 열람
								%>
										<a href="#" onClick="pop_Window('/sales/saupbu_profit_loss_report_std.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>
								<%
									ElseIf empProfitViewSI = "Y" Then	'SI, SI2만 열람
										If saupbu_tab(i) = "SI1본부" Or saupbu_tab(i) = "SI2본부" Then
								%>
											<a href="#" onClick="pop_Window('/sales/saupbu_profit_loss_report_std.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>
								<%
										Else
											Response.Write FormatNumber(sum_amt(i, 2, k)/1000, 0)
										End If
									ElseIf empProfitViewNI = "Y" Then	'NI, ICT만 열람
										If saupbu_tab(i) = "NI본부" Or saupbu_tab(i) = "ICT본부" Then
								%>
											<a href="#" onClick="pop_Window('/sales/saupbu_profit_loss_report_std.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>
								<%
										Else
											Response.Write FormatNumber(sum_amt(i, 2, k)/1000, 0)
										End If
									ElseIf saupbu_tab(i) = bonbu Then

								%>
										<a href="#" onClick="pop_Window('/sales/saupbu_profit_loss_report_std.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>
								<%
									Else
										Response.Write FormatNumber(sum_amt(i, 2, k)/1000, 0)
									End If
								Else
								%>
								<%
								'회사간 거래 제외
								'If(k < 13 And sum_amt(i, 2, k) > 0) And (saupbu_tab(i) = "회사간거래") Then
								%>
								<!--<a href="#" onClick="pop_Window('/company_deal_detail_view.asp?cost_year=<%'=cost_year%>&cost_mm=<%'=k%>','company_deal_detail_view_pop','scrollbars=yes,width=1000,height=600')">
									<%'=FormatNumber(sum_amt(i, 2, k)/1000, 0)%>
								</a>-->
								<%' 	Else %>
									<%''=FormatNumber(sum_amt(i, 2, k)/1000, 0)%>
								<%'	End If	%>
									<a href="#" onClick="pop_Window('/sales/saupbu_profit_loss_report_std.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>

								<%End If	%>
							  </td>
								<%
								Next
								%>
			              	</tr>

							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">손익</td>
								<%
								For k = 1 To 13
								%>
								<td class="right"><%=FormatNumber(sum_amt(i, 3, k)/1000, 0)%></td>
								<%
								Next
								%>
							</tr>
							<%
							Next
							%>
							<tr>
							  	<td rowspan="3" class="first" bgcolor="#CCFFFF"><strong>계</strong></td>
								<td>매출</td>
							<%
							For k = 1 To 13
							%>
								<td class="right"><%=FormatNumber(sum_amt(0, 1, k)/1000, 0)%></td>
							<%
							Next
							%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">비용</td>
							<%
							For k = 1 To 13
							%>
								<td class="right"><%=FormatNumber((sum_amt(0, 2 ,k))/1000, 0)%></td>
							<%
							Next
							%>
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">손익</td>
							<%
							For k = 1 To 13
							%>
								<td class="right"><%=FormatNumber(sum_amt(0, 3, k)/1000, 0)%></td>
							<%
							Next
							%>
			              </tr>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="/sales/saupbu_profit_loss_total_std_excel.asp?cost_year=<%=cost_year%>" class="btnType04">엑셀다운로드</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
DBConn.Close() : Set DBConn = Nothing
%>