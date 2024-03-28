<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

slip_month=Request("slip_month")
		
from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
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
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"

if owner_company = "��ü" then
	owner_company_sql = " "
  else
	owner_company_sql = " and ( owner_company = '" + owner_company + "' ) "
end if
if card_type = "��ü" then
	card_type_sql = " "
  else
	card_type_sql = " and ( card_slip.card_type = '" + card_type + "' ) "
end if

if field_check <> "total" then
	field_sql = " and ( card_slip." + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if
order_sql = " ORDER BY slip_date ASC"

sql = base_sql + owner_company_sql + card_type_sql + field_sql + order_sql
response.write(sql)
Rs.Open Sql, Dbconn, 1

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
						<colgroup>
							<col width="4%" >
							<col width="6%" >
							<col width="12%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
							<col width="2%" >
							<col width="7%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr class="style10B">
								<th class="first" scope="col">����</th>
								<th scope="col">ī����</th>
								<th scope="col">ī���ȣ</th>
								<th scope="col">�����</th>
								<th scope="col">��������</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">�ŷ�ó��</th>
								<th scope="col">�ŷ�ó����</th>
								<th scope="col">���ް���</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">�հ�ݾ�</th>
								<th scope="col">�ΰ�����������</th>
								<th scope="col">�ΰ�������</th>
								<th scope="col">��������</th>
								<th scope="col">ǰ��</th>
								<th scope="col">����</th>
								<th scope="col">�ܰ�</th>
								<th scope="col">��ǥ��</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">������ּ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						j = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						err_cnt = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1
							if rs("cost_vat") > 0 then
								vat_yn = "����"
								vat_type = "57"
							  else
								vat_yn = "�Ұ���"
								vat_type = "  "
							end if

							Sql="select * from account where account_name = '" + rs("account") + "'"
							Set rs_acc=DbConn.Execute(Sql)
							if rs_acc.eof or rs_acc.bof then
								account_code = "ERROR"
								err_cnt = err_cnt + 1
							  else
							  	account_code = rs_acc("account_code")
							end if
						%>
							<tr class="style10">
								<td class="first">3</td>
								<td><%=rs("card_type")%></td>
								<td><%=rs("card_no")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("customer_no")%></td>
								<td><%=rs("customer")%></td>
								<td></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td>&nbsp;</td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td><%=vat_yn%></td>
								<td><%=vat_type%>&nbsp;</td>
								<td><%=account_code%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
							</tr>
					  <%
							rs.movenext()
						loop
						rs.close()
						if price_sum <> ( cost_sum + cost_vat_sum ) then
							err_msg = "�ݾ�Ȯ�� ���"
						  else
						  	err_msg = " "
						end if
						%>
							<tr class="style10B">
								<th colspan="2" class="first">�Ѱ�</th>
								<th colspan="4"><%=i%>&nbsp;��</th>
								<td></td>
							  	<th><%=formatnumber(cost_sum,0)%></th>
								<th><%=formatnumber(cost_vat_sum,0)%></th>
							  	<th>0</th>
							  	<th><%=formatnumber(price_sum,0)%></th>
								<th colspan="3">
					<% if err_cnt > 0 then	%>
                    		�������� �̺з��� <%=err_cnt%> �� �ֽ��ϴ�.
					<%   else	%>
                    			&nbsp;
                    <% end if	%>					
                                </th>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td>
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

