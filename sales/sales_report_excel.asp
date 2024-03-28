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

savefilename = from_date & " ~ " & to_date & " �� ���� ����.xls"

'���� ����
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
'	Case "��ü"
'		objBuilder.Append " "
'	Case "ȸ�簣�ŷ�", "��Ÿ�����"
'		objBuilder.Append "AND sst.saupbu = '"&sales_saupbu&"' "
'	Case Else
'		objBuilder.Append "AND eomt.org_bonbu = '"&sales_saupbu&"' "
'End Select

If sales_saupbu = "��ü" Then
	'�Ҽ� �μ� ���� ���� ���� �߰�
	If empProfitViewAll = "Y" Then
		objBuilder.Append ""
	ElseIf empProfitViewSI = "Y" Then
		objBuilder.Append "AND sst.saupbu IN ('SI1����', 'SI2����') "
	ElseIf empProfitViewNI = "Y" Then
		objBuilder.Append "AND sst.saupbu IN ('NI����', 'ICT����') "
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
		<title>ȸ�� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%'=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">��������</th>
								<th scope="col">����ȸ��</th>
								<th scope="col">��������</th>
								<th scope="col">����</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">�׷�</th>
								<th scope="col">�հ�ݾ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">ǰ���</th>
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
									'If sales_saupbu = "��Ÿ�����" Or sales_saupbu = "ȸ�簣�ŷ�" Then
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