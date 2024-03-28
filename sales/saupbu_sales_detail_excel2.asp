<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
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
'==================================================
Dim cost_month, sales_saupbu, slip_month, title_line, savefilename
Dim i, rsSales, trade_no

cost_month = request("cost_month")
sales_saupbu = request("sales_saupbu")

slip_month = mid(cost_month,1,4) & "-" & mid(cost_month,5,2)

title_line = cost_month & "월 " & sales_saupbu & " 매출 세부 내역"
savefilename = title_line & ".xls"

'엑셀 다운로드 설정
Call ViewExcelType(savefilename)

i = 0

objBuilder.Append  "SELECT trade_no, sales_date, sales_company, company, emp_name, emp_no, sales_amt, cost_amt, vat_amt, sales_memo "
objBuilder.Append  "FROM saupbu_sales "
objBuilder.Append  "   WHERE saupbu ='"&sales_saupbu&"' AND substring(sales_date,1,7) = '"&slip_month&"' "
objBuilder.Append  "ORDER BY sales_date, company "

Set rsSales = DBConn.Execute(objBuilder.ToString)
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">매출일</th>
								<th scope="col">매출회사</th>
								<th scope="col">고객사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">담당자</th>
								<th scope="col">사번</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">매출내역</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsSales.EOF
							i = i + 1

							trade_no = Mid(rsSales("trade_no"), 1, 3) & "-" & Mid(rsSales("trade_no"), 4, 2) & "-" & Mid(rsSales("trade_no"), 6)
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rsSales("sales_date")%></td>
								<td><%=rsSales("sales_company")%></td>
								<td><%=rsSales("company")%></td>
								<td><%=trade_no%></td>
								<td><%=rsSales("emp_name")%></td>
								<td><%=rsSales("emp_no")%></td>
							  	<td class="right"><%=FormatNumber(rsSales("sales_amt"), 0)%></td>
							  	<td class="right"><%=FormatNumber(rsSales("cost_amt"), 0)%></td>
							  	<td class="right"><%=FormatNumber(rsSales("vat_amt"), 0)%></td>
								<td><%=rsSales("sales_memo")%></td>
							</tr>
						<%
							rsSales.MoveNext()
						Loop
						rsSales.Close() : Set rsSales = Nothing
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
DBConn.Close() : Set DBConn = Nothing
%>