<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

from_date=Request("from_date")
to_date=Request("to_date")
sign_yn=Request("sign_yn")
slip_id=Request("slip_id")
view_date=Request("view_date")
field_check=Request("field_check")
field_view=Request("field_view")

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	view_c = "total"
	view_date = "total"
end If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

savefilename = from_date + "~" + to_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If field_check = "total" Then
	field_view = ""
End If

base_sql = "select * from sales_slip "

if sign_yn = "Y" then
	sign_sql = " where sign_yn = 'Y' "
  else
	sign_sql = " where (sign_yn = 'N' or sign_yn = 'I' or sign_yn = 'C') "
end if

if view_date = "total" then
	date_sql = " "
  else
  	date_sql = "and ("&view_date&" >='"&from_date&"' and "&view_date&" <= '"&to_date&"') "
end if

if slip_id = "T" then
	slip_sql = " "
  else
	slip_sql = " and slip_id = '"&slip_id&"' "
end if

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " ORDER BY slip_no DESC"

sql = base_sql + sign_sql + date_sql + slip_sql + field_sql + order_sql
response.write(sql)
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
								<th scope="col">��ǥ����</th>
								<th scope="col">��������</th>
								<th scope="col">��ǥ��ȣ</th>
								<th scope="col">��������</th>
								<th scope="col">��꼭<br>������</th>
								<th scope="col">��꼭<br>���࿹����</th>
								<th scope="col">�ŷ�ó��</th>
								<th scope="col">�������</th>
								<th scope="col">�����Ѿ�</th>
								<th scope="col">�����Ѿ�</th>
								<th scope="col">�����Ѿ�</th>
								<th scope="col">�����Ѿ�</th>
								<th scope="col">�̼��ݻ���</th>
								<th scope="col">���ݿ�����</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
							if rs("slip_id") = "1" then
								view_slip_id = "�����ǥ"
							  else
								view_slip_id = "������ǥ"
							end if
							if rs("sign_yn") = "Y" then
								view_sign_yn = "����Ϸ�"
							  elseif rs("sign_yn") = "N" then
								view_sign_yn = "�̰���"
							  elseif rs("sign_yn") = "C" then
								view_sign_yn = "�ݷ�"
							  elseif rs("sign_yn") = "I" then
								view_sign_yn = "������"
							end if
						%>
							<tr>
								<td class="first"><%=i%></td>
							  	<td><%=view_slip_id%></td>
								<td><%=view_sign_yn%></td>
								<td><%=rs("slip_no")%>-<%=rs("slip_seq")%></td>
								<td><%=rs("sales_date")%>&nbsp;</td>
								<td>
						<% if rs("sales_yn") = "N" or rs("collect_stat") = "����" then	%>
								�̹���
                       	<%   else	%>
                        		<%=rs("bill_issue_date")%>&nbsp;
						<% end if	%>
                                </td>
								<td><%=rs("bill_due_date")%>&nbsp;</td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("margin_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("collect_tot_amt"),0)%></td>
								<td><%=rs("unpaid_memo")%></td>
								<td><%=rs("unpaid_due_date")%></td>
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

