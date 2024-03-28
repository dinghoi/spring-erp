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
'===================================================
Dim from_month, to_month, sales_saupbu, slip_month, title_line, savefilename
Dim i, j

from_month = f_Request("from_month")
to_month = f_Request("to_month")
sales_saupbu = f_Request("sales_saupbu")

title_line = from_month & "월 ~ " & to_month & "월 " & sales_saupbu & " 협업"
savefilename = title_line & ".xls"

Call ViewExcelType(savefilename)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
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
								<th class="first" scope="col">일자</th>
								<th scope="col">사업부</th>
								<th scope="col">거래처</th>
								<th scope="col">협업지원건수</th>
								<th scope="col">협업지원비용</th>
								<th scope="col">받은협업건수</th>
								<th scope="col">받은협업비용</th>
								<th scope="col">총 건수</th>
								<th scope="col">총 비용</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim std_cost, rsCowork, arrCowork, cost_month, as_company
						Dim as_give_cowork, as_get_cowork, cowork_give_cost, cowork_get_cost
						Dim as_total, cost_total

						'표준 인건비
						std_cost = 30000	'2021년도 적용

						objBuilder.Append "SELECT as_month, saupbu, as_company, as_give_cowork, as_get_cowork, "
						objBuilder.Append "	cowork_give_cost, cowork_get_cost,  "
						objBuilder.Append "	as_total, "
						objBuilder.Append "	(cowork_give_cost + cowork_get_cost) AS 'cost_total' "
						objBuilder.Append "FROM ( "
						objBuilder.Append "	SELECT as_month, trdt.saupbu, as_company, as_give_cowork, as_get_cowork, "
						objBuilder.Append "		(as_give_cowork * "&std_cost&" * -1) AS 'cowork_give_cost', "
						objBuilder.Append "		(as_get_cowork * "&std_cost&") AS 'cowork_get_cost', "
						objBuilder.Append "		(as_give_cowork + as_get_cowork) AS 'as_total' "
						objBuilder.Append "	FROM as_acpt_status AS aast "
						objBuilder.Append "	INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
						objBuilder.Append "		AND trdt.trade_id = '매출' "
						objBuilder.Append "	WHERE (aast.as_month >= '"&from_month&"' AND aast.as_month <= '"&to_month&"') "
						objBuilder.Append "		AND (as_give_cowork > 0 OR as_get_cowork > 0) "

						If sales_saupbu <> "" Then
							If sales_saupbu = "기타사업부" Then
								objBuilder.Append "		AND trdt.saupbu = '' "
							Else
								objBuilder.Append "		AND trdt.saupbu = '"&sales_saupbu&"' "
							End If
						End If
						objBuilder.Append ") r1 "
						objBuilder.append "ORDER BY as_month, saupbu, as_company "

						'Response.write objBuilder.ToString()

						Set rsCowork = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsCowork.EOF Then
							arrCowork = rsCowork.getRows()
						End If
						rsCowork.Close() : Set rsCowork = Nothing
						DBConn.Close() : Set DBConn = Nothing

						j = 0

						If IsArray(arrCowork) Then
							For i = LBound(arrCowork) To UBound(arrCowork, 2)
								cost_month = arrCowork(0, i)	'일자
								saupbu = arrCowork(1, i)	'본부명
								as_company = arrCowork(2, i)	'거래처명
								as_give_cowork = CDbl(f_toString(arrCowork(3, i), 0))	'협업 지원 건수
								as_get_cowork = CDbl(f_toString(arrCowork(4, i), 0))	'받은 지원 건수
								cowork_give_cost = CDbl(f_toString(arrCowork(5, i), 0))	'협업 지원 비용
								cowork_get_cost = CDbl(f_toString(arrCowork(6, i), 0))	'받은 협업 비용
								as_total = CDbl(f_toString(arrCowork(7, i), 0))	'총 건수
								cost_total = CDbl(f_toString(arrCowork(8, i), 0))	'총 비용

								j = j + 1
						%>
							<tr>
								<td class="first"><%=j%></td>
								<td><%=cost_month%></td>
								<td><%=saupbu%></td>
								<td><%=as_company%></td>
								<td><%=FormatNumber(as_give_cowork, 0)%></td>
								<td><%=FormatNumber(cowork_give_cost, 0)%></td>
								<td><%=FormatNumber(as_get_cowork, 0)%></td>
								<td><%=FormatNumber(cowork_get_cost, 0)%></td>
								<td><%=FormatNumber(as_total, 0)%></td>
								<td><%=FormatNumber(cost_total, 0)%></td>
							</tr>
						<%
							Next
						End If
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>