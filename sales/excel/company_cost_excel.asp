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

'Response.write sales_saupbu

title_line = from_month & "�� ~ " & to_month & "�� " & sales_saupbu & " �ŷ�ó�� ����"
savefilename = title_line & ".xls"

Call ViewExcelType(savefilename)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th class="first" scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">�ŷ�ó</th>
								<th scope="col">����</th>
								<th scope="col">����������(�ΰǺ�)</th>
								<th scope="col">����������(�Ϲݰ��)</th>
								<th scope="col">����ΰ����</th>
								<!--<th scope="col">����</th>-->
								<th scope="col">�ι������</th>
								<th scope="col">��������</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim rsCompCost, arrCompCost, company_name, sales_cost, pay_cost, general_cost, common_cost
						Dim part_cost, manage_cost, profit_cost, cost_month, cowork_cost

						objBuilder.Append "SELECT * FROM ("
						objBuilder.Append "SELECT cost_month, saupbu, company_name, "
						objBuilder.Append "	SUM(sales_cost) AS 'sales_cost', SUM(pay_cost) AS 'pay_cost', SUM(general_cost) AS 'general_cost', "
						objBuilder.Append "	SUM(common_cost) AS 'common_cost', SUM(part_cost) AS 'part_cost', SUM(manage_cost) AS 'manage_cost', "
						objBuilder.Append "	SUM(profit_cost) AS 'profit_cost', "
						objBuilder.Append "	SUM(pay_cost) + SUM(general_cost) + SUM(common_cost) + SUM(part_cost) + SUM(manage_cost)  AS 'c_cost' "
						'objBuilder.Append "	SUM(cowork_give_cost + cowork_get_cost) AS 'cowork_cost' "
						objBuilder.Append "FROM company_cost_profit "
						objBuilder.Append "WHERE (cost_month >= '"&from_month&"' AND cost_month <= '"&to_month&"') "

						If sales_saupbu <> "" Then
							objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
						End If

						objBuilder.Append "GROUP BY cost_month, saupbu, company_name "
						objBuilder.Append "ORDER BY cost_month, saupbu, company_name "
						objBuilder.Append ") r1 WHERE r1.sales_cost <> 0 OR r1.c_cost <> 0 "

						'Response.write objBuilder.ToString()

						Set rsCompCost = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsCompCost.EOF Then
							arrCompCost = rsCompCost.getRows()
						End If
						rsCompCost.Close() : Set rsCompCost = Nothing
						DBConn.Close() : Set DBConn = Nothing

						j = 0

						If IsArray(arrCompCost) Then
							For i = LBound(arrCompCost) To UBound(arrCompCost, 2)
								cost_month = arrCompCost(0, i)
								saupbu = arrCompCost(1, i)
								company_name = arrCompCost(2, i)	'�ŷ�ó��
								sales_cost = CDbl(f_toString(arrCompCost(3, i), 0))	'�ŷ�ó�� ����
								pay_cost = CDbl(f_toString(arrCompCost(4, i), 0))	'����������(�ΰǺ�)
								general_cost = CDbl(f_toString(arrCompCost(5, i), 0))	'����������(�Ϲݰ��)
								common_cost = CDbl(f_toString(arrCompCost(6, i), 0))	'����ΰ�����
								part_cost = CDbl(f_toString(arrCompCost(7, i), 0))	'�ι������
								manage_cost = CDbl(f_toString(arrCompCost(8, i), 0))	'����Ʈ�� ��������(���� ����)
								profit_cost = CDbl(f_toString(arrCompCost(9, i), 0))	'NKP ����
								'cowork_cost = CDbl(f_toString(arrCompCost(11, j), 0))	'���� ���

								j = j + 1
						%>
							<tr>
								<td class="first"><%=j%></td>
								<td><%=cost_month%></td>
								<td><%=saupbu%></td>
								<td><%=company_name%></td>
								<td><%=FormatNumber(sales_cost, 0)%></td>
								<td><%=FormatNumber(pay_cost, 0)%></td>
								<td><%=FormatNumber(general_cost, 0)%></td>
								<td><%=FormatNumber(common_cost, 0)%></td>
								<!--<td><%'=FormatNumber(cowork_cost, 0)%></td>-->
								<td><%=FormatNumber(part_cost, 0)%></td>
								<td><%=FormatNumber(manage_cost, 0)%></td>
								<td><%=FormatNumber(profit_cost, 0)%></td>
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