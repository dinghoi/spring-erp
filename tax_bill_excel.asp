<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw

from_date=Request("from_date")
to_date=Request("to_date")
bill_id=Request("bill_id")

bill_id = "����"

savefilename = from_date + "~" + to_date + " " + bill_id + " ���ݰ�꼭 ����.xls"

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

' ���Ǻ� ��ȸ.........
sql = "select * from general_cost where tax_bill_yn = 'Y' and (slip_date >= '" + from_date  + "' and slip_date <= '" + to_date  + "') ORDER BY customer, slip_gubun, slip_date ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>����ȸ��ý���</title>
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
								<th class="first" scope="col">�ŷ�ó</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">������</th>
								<th scope="col">����</th>
								<th scope="col">�׸�</th>
								<th scope="col">���೻��</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">����ȸ��</th>
                                <th scope="col">����μ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1
							if rs("end_yn") = "Y" then
								end_yn = "����"
							  else
							  	end_yn = "����"
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
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("emp_company")%></td>
								<td><%=rs("org_name")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th class="first">�Ѱ�</th>
								<th colspan="1"><%=i%>&nbsp;��</th>
							  	<th colspan="4">&nbsp</th>
							  	<th><%=formatnumber(price_sum,0)%></th>
							  	<th><%=formatnumber(cost_sum,0)%></th>
								<th><%=formatnumber(cost_vat_sum,0)%></th>
								<th colspan="2">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

