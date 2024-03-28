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
Dim sales_saupbu, field_check, field_view, sales_yymm
Dim savefilename, field_sql, rs, from_date, to_date

'Dim sales_month
'sales_month = Request("sales_month")
'sales_yymm = Mid(sales_month, 1, 4) & "-" & Mid(sales_month, 5, 2)

from_date = Request("from_date")
to_date = Request("to_date")

sales_saupbu = Request("sales_saupbu")
field_check = Request("field_check")
field_view = Request("field_view")

savefilename = from_date & " ~ " & to_date & " 월 매출 내역.xls"

'엑셀 지정
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT sst.sales_date, sst.sales_company, sst.saupbu, sst.company, sst.trade_no, "
objBuilder.Append "	sst.group_name, sst.sales_amt, sst.cost_amt, sst.vat_amt, sst.emp_name, sst.sales_memo, "
objBuilder.Append "	sst.approve_no,	sst.emp_no	"
objBuilder.Append "FROM saupbu_sales AS sst "
objBuilder.Append "WHERE sales_date BETWEEN '"&from_date&"' AND '"&to_date&"' "

If field_check <> "total" Then
	objBuilder.Append "AND "&field_check&" LIKE '%"&field_view&"%' "
End If

'Select Case sales_saupbu
'	Case "전체"
'		objBuilder.Append " "
'	Case "회사간거래", "기타사업부"
'		objBuilder.Append "AND sst.saupbu = '"&sales_saupbu&"' "
'	Case Else
'		objBuilder.Append "AND eomt.org_bonbu = '"&sales_saupbu&"' "
'End Select

If sales_saupbu = "전체" Then
	'소속 부서 제한 열람 조건 추가
	If empProfitViewAll = "Y" Then
		objBuilder.Append ""
	ElseIf empProfitViewSI = "Y" Then
		objBuilder.Append "AND sst.saupbu IN ('SI1본부', 'SI2본부') "
	ElseIf empProfitViewNI = "Y" Then
		objBuilder.Append "AND sst.saupbu IN ('NI본부', 'ICT본부') "
	Else
		objBuilder.Append "AND sst.saupbu = '"&bonbu&"' "
	End If
Else
	objBuilder.Append "AND sst.saupbu = '"&sales_saupbu&"' "
End If

objBuilder.Append "ORDER BY sst.sales_date, sst.saupbu ASC "

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>회계 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%'=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">매출일자</th>
								<th scope="col">매출회사</th>
								<th scope="col">영업본부</th>
								<th scope="col">고객사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">그룹</th>
								<th scope="col">합계금액</th>
								<th scope="col">공급가액</th>
								<th scope="col">세액</th>
								<th scope="col">담당자</th>
								<th scope="col">품목명</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rs.EOF
						%>
							<tr>
								<td class="first"><%=rs("sales_date")%></td>
								<td><%=rs("sales_company")%></td>
								<td>
								<%
									'If sales_saupbu = "기타사업부" Or sales_saupbu = "회사간거래" Then
										Response.Write rs("saupbu")
									'Else
									'	Response.Write rs("org_bonbu")
									'End If
								%>
								</td>
								<td><%=rs("company")%></td>
								<td><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=FormatNumber(rs("cost_amt"),0)%></td>
								<td class="right"><%=FormatNumber(rs("vat_amt"),0)%></td>
								<td><%=rs("emp_name")%>&nbsp;</td>
								<td class="left"><%=rs("sales_memo")%></td>
							</tr>
						<%
							rs.MoveNext()
						Loop
						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>