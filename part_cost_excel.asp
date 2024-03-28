<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'on Error resume next

Dim from_date
Dim to_date
Dim win_sw

cost_month=Request("cost_month")
sales_saupbu=Request("sales_saupbu")

if cost_month = "" then
	before_date = dateadd("m",-1,now())
	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
	sales_saupbu = "��ü"
end If

if sales_saupbu = "��ü" then
	condi_sql = ""
  else
  	condi_sql = " and saupbu ='"&sales_saupbu&"'"
end if
mm = mid(cost_month,5,2)
cost_year = mid(cost_month,1,4)

sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '�ι������'"
Set rs=DbConn.Execute(SQL)
tot_part_cost = clng(rs("tot_cost"))
rs.close()

sql = "select * from company_as where (as_month = '"&cost_month&"')"&condi_sql&"  order by as_company"
rs.Open sql, Dbconn, 1

title_line = cost_year + "��" + mm + "�� " + sales_saupbu + " �κ� ����� �����Ȳ"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

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
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">ȸ��</th>
								<th scope="col">�����</th>
								<th scope="col">���ݰǼ�</th>
								<th scope="col">���ݿ�</th>
								<th scope="col">������(%)</th>
								<th scope="col">�����ݾ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						remote_sum = 0
						visit_sum = 0
						charge_per_sum = 0
						charge_cost_sum = 0

						do until rs.eof
							charge_cost = int(rs("charge_per") * tot_part_cost)
							remote_sum = rs("remote_cnt") + remote_sum
							visit_sum = rs("visit_cnt") + visit_sum
							charge_per_sum = rs("charge_per") + charge_per_sum
							charge_cost_sum = rs("cost_amt") + charge_cost_sum
						%>
							<tr>
								<td class="first"><%=rs("as_company")%></td>
								<td><%=rs("saupbu")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("remote_cnt"),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("visit_cnt"),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("charge_per"),5)%>&nbsp;%&nbsp;</td>
								<td class="right"><%=formatnumber(rs("cost_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						%>
							<tr>
								<td bgcolor="#FFE8E8" class="first">�Ѱ�</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(remote_sum,0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(visit_sum,0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(charge_per_sum,5)%>&nbsp;%&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(charge_cost_sum,0)%>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				<br>
		</div>
	</div>
	</body>
</html>

