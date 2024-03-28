<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
		 
cost_month=Request("cost_month")
	
from_date = mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
	
be_yy = int(mid(cost_month,1,4))
be_mm = int(mid(cost_month,5) - 1)
if be_mm = 0 then
	be_month = cstr(be_yy - 1) + "12"
  else
	be_month = cstr(be_yy) + right("0" + cstr(be_mm),2)
end if

title_line = mid(cost_month,1,4) + "��" + mid(cost_month,5) + "�� ȸ�� ���� ������� ��� �� ���ݾ�"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

sql = "select car_no, car_name, oil_kind, mg_ce_id, user_name, team, org_name, sum(far) as sum_far, sum(oil_price) as sum_price, sum(repair_cost) as sum_repair  from transit_cost where (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cancel_yn = 'N') and car_owner = 'ȸ��' group by car_no, mg_ce_id order by car_no, user_name "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" border="1" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="10%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
							  <th colspan="3" class="first" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� ����</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">������ ����</th>
							  <th rowspan="2" scope="col">����Ÿ�</th>
							  <th colspan="4" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� �ݾ�</th>
							  <th colspan="4" style=" border-bottom:1px solid #e3e3e3;" scope="col">�� ���ݾ�</th>
							  <th rowspan="2" scope="col">����</th>
						  </tr>
							<tr>
							  <th class="first" scope="col">������ȣ</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">����</th>
							  <th scope="col">����</th>
							  <th scope="col">������</th>
							  <th scope="col">������</th>
							  <th scope="col">�ܰ�</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">�Ҹ�ǰ</th>
							  <th scope="col">�Ұ�</th>
							  <th scope="col">����������</th>
							  <th scope="col">����ī��</th>
							  <th scope="col">������</th>
							  <th scope="col">�Ұ�</th>
						  </tr>
						</thead>
						<tbody>
					<%
						rs.Open sql, Dbconn, 1
						do until rs.eof
							if rs("team") = "������" or rs("team") = "SM1��" or rs("team") = "Repair��" or rs("team") = "SM2��" then
								oil_unit_id = "1"
							  else
								oil_unit_id = "2"
							end if
							
							sql = "select * from oil_unit where oil_unit_month = '"&cost_month&"' and oil_unit_id = '"&oil_unit_id&"' and oil_kind = '"&rs("oil_kind")&"'"
							set rs_etc=dbconn.execute(sql)
							if rs_etc.eof or rs_etc.bof then
								oil_unit = 1
							  else
								oil_unit = rs_etc("oil_unit_average")
							end if								
							rs_etc.close()

							if oil_kind = "����" then
								oil_cost = round(cdbl(rs("sum_far")) * oil_unit / 7)
							  else
								oil_cost = round(cdbl(rs("sum_far")) * oil_unit / 10)
							end if
							somopum = cdbl(rs("sum_far")) * 25
							
' ���� ī����
							juyoo_card_price = 0
							sql = "select count(*) as c_cnt,sum(price) as price from card_slip where (emp_no='"&user_id&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and card_type like '%����%'"
								
							Set rs_etc = Dbconn.Execute (sql)
							if cint(rs_etc("c_cnt")) <>  0 then
								juyoo_card_price = cdbl(rs_etc("price"))
							end if
							rs_etc.close()
					%>
							<tr>
								<td height="25" class="first"><%=rs("car_no")%></td>
								<td><%=rs("car_name")%></td>
								<td><%=rs("oil_kind")%></td>
								<td><%=rs("user_name")%></td>
								<td><%=rs("org_name")%></td>
								<td class="right"><%=formatnumber(rs("sum_far"),0)%></td>
								<td class="right"><%=formatnumber(oil_unit,0)%></td>
								<td class="right"><%=formatnumber(oil_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum,0)%></td>
								<td class="right"><%=formatnumber(oil_cost + somopum,0)%></td>
								<td class="right"><%=formatnumber(rs("sum_price"),0)%></td>
								<td class="right"><%=formatnumber(juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(rs("sum_repair"),0)%></td>
								<td class="right"><%=formatnumber(cdbl(rs("sum_price")) + juyoo_card_price + cdbl(rs("sum_repair")),0)%></td>
								<td class="right"><%=formatnumber(cdbl(rs("sum_price")) + juyoo_card_price + cdbl(rs("sum_repair")) - oil_cost - somopum,0)%></td>
							</tr>
					<%
							rs.movenext()
						loop
					%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

