<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

cost_month = request("cost_month")

slip_month = mid(cost_month,1,4) + "-" + mid(cost_month,5,2)

title_line = cost_month + "�� ���ݰ�꼭 ����"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from general_cost where (tax_bill_yn = 'Y') and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
Rs.Open Sql, Dbconn, 1

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
								<th scope="col">���ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��</th>
								<th scope="col">������</th>
								<th scope="col">����ó</th>
								<th scope="col">�����</th>
								<th scope="col">��������</th>
								<th scope="col">�������</th>
								<th scope="col">����</th>
								<th scope="col">��翵�������</th>
								<th scope="col">���־�ü</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">�������</th>
								<th scope="col">��뱸��</th>
								<th scope="col">��������</th>
								<th scope="col">���೻��</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("emp_company")%></td>
								<td><%=rs("bonbu")%></td>
								<td><%=rs("saupbu")%></td>
								<td><%=rs("team")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("reside_place")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("slip_seq")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("mg_saupbu")%></td>
								<td><%=rs("customer")%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("cost_center")%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
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

