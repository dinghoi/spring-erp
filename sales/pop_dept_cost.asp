<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
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
Dim dept, dt
Dim title_line
Dim totalYn : totalYn = "Y"
Dim rsSales
Dim tot_cost_amt, tot_charge_per, tot_company_cost
Dim salesDate

dept = Request("dept")
dt = Request("dt")

title_line = "사업부내 고객사별 매출액 비율"
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
	</head>
<body>
	<div style="margin:0px 10px 0px 10px;">
		<div id="container">
		<h3 class="stit">* <%=title_line%></h3>
			<table cellpadding="0" cellspacing="0" summary="" class="tableList">
			<colgroup>
				<col width="20%" >
				<col width="20%" >
				<col width="*" >
				<col width="20%" >
			</colgroup>
			<thead>
				<tr>
					<th class="first" scope="col">회사</th>
					<th scope="col">사업부</th>
					<th scope="col">고객사</th>
					<th scope="col">매출</th>
				</tr>
			</thead>
			<tbody>
			<%
			tot_cost_amt = 0
			tot_charge_per = 0
			tot_company_cost = 0

			salesDate = Left(dt, 4) & "-" & Right(dt, 2)

			objBuilder.Append "SELECT sales_company, saupbu, company, SUM(cost_amt) AS cost_amt "
			objBuilder.Append "FROM saupbu_sales "
			objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&salesDate&"' "
			objBuilder.Append "AND saupbu ='"&dept&"' "
			objBuilder.Append "GROUP BY sales_company, saupbu ,company "

			Set rsSales = Server.CreateObject("ADODB.RecordSet")
			rsSales.Open objBuilder.ToString(), DBConn, 1
			objBuilder.Clear()

			If rsSales.BOF Or rsSales.EOF Then
				totalYn = "N"
			%>
				<tr>
					<td class="first" colspan="4">해당 데이터가 없습니다.</td>
				</tr>
			<%
			Else
				Do Until rsSales.EOF
					tot_cost_amt = tot_cost_amt + rsSales("cost_amt")
					%>
					<tr>
						<td class="first"><%=rsSales("sales_company")%></td>
						<td><%=rsSales("saupbu")%></td>
						<td><%=rsSales("company")%>&nbsp;</td>
						<td class="right"><%=FormatNumber(rsSales("cost_amt"),0)%>&nbsp;</td>
					</tr>
					<%
					rsSales.MoveNext()
				Loop
			End If

			rsSales.Close()
			Set rsSales = Nothing

			DBConn.Close()
			Set DBConn = Nothing

			If totalYn = "Y" Then
			%>
				<tr bgcolor="#FFE8E8">
					<td class="first" colspan="3">계</td>
					<td class="right"><%=FormatNumber(tot_cost_amt, 0)%>&nbsp;</td>
				</tr>
			<%End If%>
			</tbody>
			</table>
		</div>
	</div>
</body>
</html>
