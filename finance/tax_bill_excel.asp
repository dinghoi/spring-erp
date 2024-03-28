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
Dim from_date, to_date, bill_id, savefilename
Dim rs, title_line

from_date = Request("from_date")
to_date = Request("to_date")
bill_id = Request("bill_id")

bill_id = "매입"

savefilename = from_date & "~" & to_date & " " & bill_id & " 세금계산서 내역.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

' 조건별 조회
'sql = "select * from general_cost where tax_bill_yn = 'Y' and (slip_date >= '" + from_date  + "' and slip_date <= '" + to_date  + "') ORDER BY customer, slip_gubun, slip_date ASC"
objBuilder.Append "SELECT glct.customer, glct.customer_no, glct.slip_date, glct.slip_gubun, "
objBuilder.Append "	glct.account_item, glct.slip_memo, glct.price, glct.cost, glct.cost_vat, "
objbuilder.Append "	glct.emp_company, glct.org_name, glct.emp_no, glct.emp_name, glct.end_yn, "
objBuilder.Append "	eomt.org_company, eomt.org_name AS emp_org_name "
objBuilder.Append "FROM general_cost AS glct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON glct.emp_no = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE tax_bill_yn = 'Y' "
objBuilder.Append "	AND (slip_date >= '" + from_date  + "' AND slip_date <= '" + to_date  + "') "
objBuilder.Append "ORDER BY customer, slip_gubun, slip_date "

Set rs = Server.CreateObject("ADODB.Recordset")
Rs.Open objBuilder.ToString(), DBConn, 1
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처</th>
								<th scope="col">사업자번호</th>
								<th scope="col">발행일</th>
								<th scope="col">유형</th>
								<th scope="col">항목</th>
								<th scope="col">발행내역</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">발행회사</th>
                                <th scope="col">발행부서</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i, price_sum, cost_sum, cost_vat_sum
						Dim end_yn, customer_no

						i = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0

						Do Until rs.EOF
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1
							if rs("end_yn") = "Y" then
								end_yn = "마감"
							  else
							  	end_yn = "진행"
							end if
							customer_no = mid(rs("customer_no"),1,3) + "-" + mid(rs("customer_no"),4,2) + "-" + right(rs("customer_no"),5)
						%>
							<tr>
								<td class="first"><%=rs("customer")%></td>
								<td><%=customer_no%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("slip_memo")%></td>
							  	<td class="right"><%=FormatNumber(rs("price"), 0)%></td>
							  	<td class="right"><%=FormatNumber(rs("cost"), 0)%></td>
							  	<td class="right"><%=FormatNumber(rs("cost_vat"), 0)%></td>
								<td><%=rs("emp_company")%></td>
								<td><%=rs("emp_org_name")%></td>
							</tr>
						<%
							rs.MoveNext()
						Loop

						rs.close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
							<tr>
								<th class="first">총계</th>
								<th colspan="1"><%=i%>&nbsp;건</th>
							  	<th colspan="4">&nbsp</th>
							  	<th><%=FormatNumber(price_sum, 0)%></th>
							  	<th><%=FormatNumber(cost_sum, 0)%></th>
								<th><%=FormatNumber(cost_vat_sum, 0)%></th>
								<th colspan="2">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>

