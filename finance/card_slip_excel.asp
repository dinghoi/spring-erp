<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

slip_month = Request("slip_month")

from_date = mid(slip_month,1,4) & "-" & mid(slip_month,5,2) & "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
owner_company = Request("owner_company")
card_type = Request("card_type")
field_check = Request("field_check")
field_view = Request("field_view")

title_line = slip_month + " ī�� ��ǥ ����"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'base_sql = "select * from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
base_sql = "SELECT crst.emp_no, crst.emp_name, crst.card_type, crst.card_no, crst.slip_date, "
base_sql = base_sql & "	crst.cost_center, crst.account, crst.account_item, crst.pl_yn, "
base_sql = base_sql & "	crst.price, crst.cost, crst.cost_vat, "
base_sql = base_sql & "	crst.emp_company, crst.bonbu, crst.saupbu, crst.team, crst.org_name, "
base_sql = base_sql & "	crst.reside_place, crst.reside_company, emtt.cost_center AS costCenter, "
base_sql = base_sql & "	eomt.org_name AS orgName, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, "
base_sql = base_sql & "	eomt.org_team, eomt.org_reside_place, eomt.org_reside_company "
base_sql = base_sql & "FROM card_slip AS crst "
base_sql = base_sql & "INNER JOIN emp_master AS emtt ON crst.emp_no = emtt.emp_no "
base_sql = base_sql & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
base_sql = base_sql & "WHERE (crst.slip_date >='"&from_date&"' AND crst.slip_date <='"&to_date&"') "

if owner_company = "��ü" then
	owner_company_sql = " "
  else
	owner_company_sql = " and owner_company = '" & owner_company & "' "
end if
if card_type = "��ü" then
	card_type_sql = " "
  else
	card_type_sql = " and crst.card_type = '" & card_type & "' "
end if

if field_check <> "total" then
	field_sql = " and crst." & field_check & " like '%" & field_view & "%' "
  else
  	field_sql = " "
end if
order_sql = " ORDER BY crst.slip_date ASC"

SQL = base_sql & owner_company_sql & card_type_sql & field_sql & order_sql

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open SQL, Dbconn, 1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
	<style type="text/css">
    <!--
    	.style10 {font-size: 10px; font-family: "����ü", "����ü", Seoul; }
        .style10B {font-size: 10px; font-weight: bold; font-family: "����ü", "����ü", Seoul; }
    -->
    </style>
		<title>���� ȸ�� �ý���</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr class="style10B">
								<th class="first" scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��</th>
								<th scope="col">������</th>
								<th scope="col">����ó</th>
								<th scope="col">����ȸ��</th>
								<th scope="col">�����</th>
								<th scope="col">ī����</th>
								<th scope="col">ī���ȣ</th>
								<th scope="col">��������</th>
								<th scope="col">�������</th>
								<th scope="col">��������</th>
								<th scope="col">�׸�</th>
								<th scope="col">���ް���</th>
								<th scope="col">����</th>
								<th scope="col">�հ�ݾ�</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
						%>
							<tr class="style10">
								<td><%=rs("org_company")%></td>
								<td><%=rs("org_bonbu")%></td>
								<td><%=rs("org_saupbu")%></td>
								<td><%=rs("org_team")%></td>
								<td><%=rs("orgName")%></td>
								<td><%=rs("org_reside_place")%></td>
								<td><%=rs("org_reside_company")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("card_type")%></td>
								<td><%=rs("card_no")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("costCenter")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("account_item")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td><%=rs("pl_yn")%></td>
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

