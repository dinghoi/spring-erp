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
Dim bill_month, owner_company, field_check, field_view
Dim from_date, end_date, to_date, savefilename
Dim rsTax, arrTax, title_line

bill_month = f_Request("bill_month")
owner_company = f_Request("owner_company")
field_check = f_Request("field_check")
field_view = f_Request("field_view")

from_date = Mid(bill_month, 1, 4) & "-" & Mid(bill_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

title_line = bill_month & "월 이세로 세금계산서 내역"

savefilename = title_line & ".xls"

'엑셀 지정
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT r1.bill_date, r1.owner_company, r1.trade_no, r1.trade_name, r1.trade_owner, "
objBuilder.Append "	r1.price, r1.cost, r1.cost_vat, r1.bill_collect, emtt.emp_name, r1.receive_email, r1.tax_bill_memo "
objBuilder.Append "FROM ("
objBuilder.Append "SELECT tabt.receive_email, tabt.trade_no, tabt.bill_date, tabt.trade_owner, "
objBuilder.Append "	tabt.owner_company, tabt.trade_name, tabt.price, tabt.cost, tabt.cost_vat, "
objBuilder.Append "	tabt.bill_collect, tabt.tax_bill_memo, tabt.approve_no, "
objBuilder.Append "	CASE WHEN trat.trade_code = '' OR IFNULL(trat.trade_code, '') = '' THEN 'N' ELSE 'Y' END AS 'trade_sw', "

objBuilder.Append "	IF(receive_email = '' OR IFNULL(receive_email, '') = '', NULL, "
objBuilder.Append "		(SELECT emtt.emp_no FROM emp_master AS emtt "
objBuilder.Append "		WHERE emtt.emp_email = SUBSTRING(tabt.receive_email, 1, INSTR(tabt.receive_email, '@') - 1) "
objBuilder.Append "			AND emtt.emp_pay_id <> '2' "
objBuilder.Append "		LIMIT 1	"
objBuilder.Append "	)) AS 'emp_no' "
objBuilder.Append "FROM tax_bill AS tabt "
objBuilder.Append "LEFT OUTER JOIN trade AS trat ON tabt.trade_no = trat.trade_no "
objBuilder.Append "WHERE (tabt.bill_date >='"&from_date&"' AND tabt.bill_date <='"&to_date&"') "
objBuilder.Append "	AND tabt.end_yn = 'Y' AND tabt.cost_reg_yn = 'N' AND tabt.bill_id ='1' "
objBuilder.Append ") r1 "
objBuilder.Append "LEFT OUTER JOIN memb AS memt ON r1.emp_no = memt.user_id "
objBuilder.Append "	AND memt.grade < '5' "
objBuilder.Append "LEFT OUTER JOIN emp_master AS emtt ON r1.emp_no = emtt.emp_no "
objBuilder.Append "WHERE 1=1 "

If field_check <> "total" Then
	objBuilder.Append "AND "&field_check&" LIKE '%"&field_view&"%' "
End If

If owner_company <> "전체" Then
	objBuilder.Append "AND owner_company = '"&owner_company&"' "
End If

objBuilder.Append "ORDER BY bill_date, approve_no ASC "

Set rsTax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsTax.EOF Then
	arrTax = rsTax.getRows()
End If

rsTax.Close() : Set rsTax = Nothing
DBConn.Close() : Set DBConn = Nothing
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
						<colgroup>
							<col width="6%" >
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="6%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">발행일</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호명</th>
								<th scope="col">대표자명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">청구</th>
								<th scope="col">담당자</th>
								<th scope="col">공급받는자이메일</th>
								<th scope="col">거래내역</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i, t_bill_date, t_owner_company, t_trade_no, t_trade_name, t_trade_owner
						Dim t_price, t_cost, t_cost_vat, t_bill_collect, t_emp_name, t_receive_email, t_tax_bill_memo

						If IsArray(arrTax) Then
							For i = LBound(arrTax) To UBound(arrTax, 2)
								t_bill_date = arrTax(0, i)
								t_owner_company = arrTax(1, i)
								t_trade_no = arrTax(2, i)
								t_trade_name = arrTax(3, i)
								t_trade_owner = arrTax(4, i)
								t_price = arrTax(5, i)
								t_cost = arrTax(6, i)
								t_cost_vat = arrTax(7, i)
								t_bill_collect = arrTax(8, i)
								t_emp_name = arrTax(9, i)
								t_receive_email = arrTax(10, i)
								t_tax_bill_memo = arrTax(11, i)
						%>
							<tr>
								<td class="first"><%=t_bill_date%></td>
								<td><%=t_owner_company%></td>
								<td><%=Mid(t_trade_no, 1, 3)%>-<%=Mid(t_trade_no, 4, 2)%>-<%=Right(t_trade_no, 5)%></td>
								<td><%=t_trade_name%></td>
								<td><%=t_trade_owner%></td>
								<td class="right"><%=FormatNumber(t_price, 0)%></td>
								<td class="right"><%=FormatNumber(t_cost, 0)%></td>
								<td class="right"><%=FormatNumber(t_cost_vat, 0)%></td>
								<td><%=t_bill_collect%></td>
								<td><%=t_emp_name%></td>
								<td><%=t_receive_email%></td>
								<td class="left"><%=t_tax_bill_memo%></td>
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