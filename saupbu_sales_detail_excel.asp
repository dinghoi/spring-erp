<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

cost_month = request("cost_month")
sales_saupbu = request("sales_saupbu")

slip_month = mid(cost_month,1,4) + "-" + mid(cost_month,5,2)

title_line = cost_month + "�� " + sales_saupbu + " ���� ���� ����"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

i = 0
sql = "select * from saupbu_sales where saupbu ='"&sales_saupbu&"' and substring(sales_date,1,7) = '"&slip_month&"' ORDER BY sales_date,sales_seq,company"
Response.write sql
rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
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
								<th scope="col">������</th>
								<th scope="col">����</th>
								<th scope="col">����ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">�����</th>
								<th scope="col">���</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">���⳻��</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							i = i + 1
							trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6)
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("sales_seq")%></td>
								<td><%=rs("sales_company")%></td>
								<td><%=rs("company")%></td>
								<td><%=trade_no%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("emp_no")%></td>
							  	<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_amt"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("vat_amt"),0)%></td>
								<td><%=rs("sales_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
	</body>
</html>

