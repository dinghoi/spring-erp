<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
cost_month=Request("cost_month")
team=Request("team")
saupbu=Request("saupbu")
from_date = mid(cost_month,1,4) + "-" + mid(cost_month,6,2) + "-31"

savefilename = cost_month + "�� ���� ���� ��Ȳ.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

'if position = "����" and cost_grade <> "0" then
'	sql = "select * from emp_master where emp_team = '"&team&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') order by emp_team, emp_name"
'  else
'	if saupbu = "����" then
'		sql = "select * from emp_master where emp_no = '"&user_id&"'"
'	  else
'		sql = "select * from emp_master where emp_saupbu = '"&saupbu&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') order by emp_team, emp_name"
'	end if
'end if 

if position = "����" and cost_grade > "2" then
	sql = "select * from emp_master where emp_team = '"&team&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') order by emp_team, emp_name"
  else
	if saupbu = "����" then
		sql = "select * from emp_master where emp_no = '"&user_id&"'"
	  else
		sql = "select * from emp_master where emp_saupbu = '"&saupbu&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') order by emp_team, emp_name"
	end if
end if 

rs_emp.Open sql, Dbconn, 1
	
title_line = cost_month + "�� ������ ���κ� ����� �� ������Ȳ"
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
								<th rowspan="3" class="first" scope="col">��</th>
								<th rowspan="3" scope="col">�����</th>
								<th rowspan="3" scope="col">����</th>
								<th style=" border-bottom:1px solid #e3e3e3;" scope="col">��Ư��</th>
								<th colspan="9" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� ���</th>
								<th rowspan="3" scope="col">����ī��</th>
								<th rowspan="3" scope="col">����ݾ�</th>
								<th rowspan="3" scope="col"><p>����</p><p>������</p></th>
								<th rowspan="3" scope="col">����ī��</th>
								<th rowspan="3" scope="col">��� ��</th>
							</tr>
							<tr>
							  <th style=" border-bottom:1px solid #e3e3e3;border-left:1px solid #e3e3e3;" scope="col">��û�ݾ�</th>
							  <th scope="col" style=" border-bottom:1px solid #e3e3e3;">�Ϲݺ��</th>
							  <th style=" border-bottom:1px solid #e3e3e3;" scope="col">���߱���</th>
							  <th colspan="3" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� ���� ������</th>
							  <th style=" border-bottom:1px solid #e3e3e3;" scope="col">ȸ������</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� ������</th>
							  <th rowspan="2" scope="col"><p>���ݻ��</p><p>�Ұ�</p></th>
						  </tr>
							<tr>
							  <th style="border-left:1px solid #e3e3e3;" scope="col">�ݾ�</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">����(KM)</th>
							  <th scope="col">������</th>
							  <th scope="col">�Ҹ�ǰ</th>
							  <th scope="col">������</th>
							  <th scope="col">������</th>
							  <th scope="col">�����</th>
						  </tr>
						</thead>
						<tbody>
						<%
						sum_general_cnt = 0 
						sum_general_cost = 0 
						sum_overtime_cnt = 0	 
						sum_overtime_cost = 0
						sum_fare_cost = 0	 
						sum_tot_km = 0
						sum_tot_cost = 0
						sum_somopum_cost = 0
						sum_oil_cash_cost = 0
						sum_parking_cost = 0
						sum_toll_cost = 0
						sum_juyoo_card_price = 0
						sum_card_cost = 0
						sum_cash_tot_cost = 0
						sum_return_cash = 0
						sum_repair_cost = 0
						sum_cost_sum = 0

						tot_general_cnt = 0 
						tot_general_cost = 0 
						tot_overtime_cnt = 0	 
						tot_overtime_cost = 0
						tot_fare_cost = 0	 
						tot_tot_km = 0
						tot_tot_cost = 0
						tot_somopum_cost = 0
						tot_oil_cash_cost = 0
						tot_parking_cost = 0
						tot_toll_cost = 0
						tot_juyoo_card_price = 0
						tot_card_cost = 0
						tot_cash_tot_cost = 0
						tot_return_cash = 0
						tot_cost_sum = 0
						if rs_emp.eof or rs_emp.bof then
							bi_team = ""
					      else						  
							if isnull(rs_emp("emp_team")) or rs_emp("emp_team") = "" then	
								bi_team = ""
							  else
								bi_team = rs_emp("emp_team")
							end if
						end if
							
						do until rs_emp.eof
							if isnull(rs_emp("emp_team")) or rs_emp("emp_team") = "" then
								emp_team = ""
							  else
							  	emp_team = rs_emp("emp_team")
							end if
							
							if bi_team <> emp_team then
						%>
							<tr>
								<td colspan="3" bgcolor="#EEFFFF" class="first">�Ұ�</td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_general_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_fare_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_km,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_somopum_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_oil_cash_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_parking_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_toll_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cash_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_juyoo_card_price,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_return_cash,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_repair_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_card_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cost_sum,0)%></td>
							</tr>
                        <%
								sum_general_cnt = 0 
								sum_general_cost = 0 
								sum_overtime_cnt = 0	 
								sum_overtime_cost = 0
								sum_fare_cost = 0	 
								sum_tot_km = 0
								sum_tot_cost = 0
								sum_somopum_cost = 0
								sum_oil_cash_cost = 0
								sum_parking_cost = 0
								sum_toll_cost = 0
								sum_juyoo_card_price = 0
								sum_card_cost = 0
								sum_cash_tot_cost = 0
								sum_return_cash = 0
								sum_repair_cost = 0
								sum_cost_sum = 0
								bi_team = emp_team
							end if														

							sql = "select * from person_cost where cost_month = '"&cost_month&"' and emp_no = '"&rs_emp("emp_no")&"'"
							set rs=dbconn.execute(sql)
							if rs.eof or rs.bof then
								general_cnt = 0 
								general_cost = 0 
								overtime_cnt = 0	 
								overtime_cost = 0
								gas_km = 0
								gas_unit = 0
								gas_cost = 0
								gasol_km = 0
								gasol_unit = 0 
								gasol_cost = 0	 
								diesel_km = 0
								diesel_unit = 0
								diesel_cost = 0
								somopum_cost = 0
								fare_cost = 0	 
								oil_cash_cost = 0
								repair_cost = 0
								parking_cost = 0
								toll_cost = 0
								juyoo_card_cost = 0
								juyoo_card_cost_vat = 0
								card_cost = 0
								card_cost_vat = 0
								return_cash = 0
								repair_cost = 0
								tot_km = gas_km + diesel_km + gasol_km
								tot_cost = gas_cost + diesel_cost + gasol_cost
								juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat
								card_price = card_cost + card_cost_vat
								cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
								cost_sum = 0
								car_owner = "."
							  else
								emp_end = rs("emp_end")
								car_owner = rs("car_owner")
								if car_owner = "����" then
									car_owner = "."
								end if
								general_cnt = rs("general_cnt")	 
								general_cost = rs("general_cost")	 
								overtime_cnt = rs("overtime_cnt")	 
								overtime_cost = rs("overtime_cost")	 
								gas_km = rs("gas_km")	 
								gas_unit = rs("gas_unit")	 
								gas_cost = rs("gas_cost")	 
								gasol_km = rs("gasol_km")	 
								gasol_unit = rs("gasol_unit")	 
								gasol_cost = rs("gasol_cost")	 
								diesel_km = rs("diesel_km")	 
								diesel_unit = rs("diesel_unit")	 
								diesel_cost = rs("diesel_cost")	 
								somopum_cost = rs("somopum_cost")	 
								fare_cost = rs("fare_cost")	 		 
								oil_cash_cost = rs("oil_cash_cost")	 
								repair_cost = rs("repair_cost")	 
								parking_cost = rs("parking_cost")	 
								toll_cost = rs("toll_cost")	 
								juyoo_card_cost = rs("juyoo_card_cost")	 
								juyoo_card_cost_vat = rs("juyoo_card_cost_vat")	 
								juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat
								card_cost = rs("card_cost")	 
								card_cost_vat = rs("card_cost_vat")	 
								return_cash = rs("return_cash")	 
								repair_cost = rs("repair_cost")	 
								tot_km = gas_km + diesel_km + gasol_km
								tot_cost = gas_cost + diesel_cost + gasol_cost
								cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
								if rs("car_owner") = "����" then
									cost_sum = cash_tot_cost + card_cost + overtime_cost
								  else
								  	cost_sum = cash_tot_cost + repair_cost + card_cost + overtime_cost
								end if
							end if

							sum_general_cnt = sum_general_cnt + general_cnt 
							sum_general_cost = sum_general_cost + general_cost 
							sum_overtime_cnt = sum_overtime_cnt + overtime_cnt	 
							sum_overtime_cost = sum_overtime_cost + overtime_cost
							sum_fare_cost = sum_fare_cost + fare_cost	 
							sum_tot_km = sum_tot_km + tot_km
							sum_tot_cost = sum_tot_cost + tot_cost
							sum_somopum_cost = sum_somopum_cost + somopum_cost
							sum_oil_cash_cost = sum_oil_cash_cost + oil_cash_cost
							sum_parking_cost = sum_parking_cost + parking_cost
							sum_toll_cost = sum_toll_cost + toll_cost
							sum_juyoo_card_price = sum_juyoo_card_price + juyoo_card_price
							sum_card_cost = sum_card_cost + card_cost
							sum_cash_tot_cost = sum_cash_tot_cost + cash_tot_cost
							sum_return_cash = sum_return_cash + return_cash
							sum_repair_cost = sum_repair_cost + repair_cost
							sum_cost_sum = sum_cost_sum + cost_sum

							tot_general_cnt = tot_general_cnt + general_cnt 
							tot_general_cost = tot_general_cost + general_cost 
							tot_overtime_cnt = tot_overtime_cnt + overtime_cnt	 
							tot_overtime_cost = tot_overtime_cost + overtime_cost
							tot_fare_cost = tot_fare_cost + fare_cost	 
							tot_tot_km = tot_tot_km + tot_km
							tot_tot_cost = tot_tot_cost + tot_cost
							tot_somopum_cost = tot_somopum_cost + somopum_cost
							tot_oil_cash_cost = tot_oil_cash_cost + oil_cash_cost
							tot_parking_cost = tot_parking_cost + parking_cost
							tot_toll_cost = tot_toll_cost + toll_cost
							tot_juyoo_card_price = tot_juyoo_card_price + juyoo_card_price
							tot_card_cost = tot_card_cost + card_cost
							tot_cash_tot_cost = tot_cash_tot_cost + cash_tot_cost
							tot_return_cash = tot_return_cash + return_cash
							tot_repair_cost = tot_repair_cost + repair_cost
							tot_cost_sum = tot_cost_sum + cost_sum
'							if cost_sum <> 0 then
							
							if emp_end = "�ٹ�" then
								emp_view = rs_emp("emp_name") + " " + rs_emp("emp_job")
							  else
							  	emp_view = "��� " + rs_emp("emp_name")
							end if
							if emp_end = "�ٹ�" or ( emp_end = "���" and juyoo_card_price > 0 ) then
						%>
							<tr>
								<td class="first"><%=rs_emp("emp_team")%>&nbsp;</td>
						<% if emp_end = "���" then	%>
								<td bgcolor="#FFCCFF"><%=emp_view%></td>
                        <%   else	%>
								<td><%=emp_view%></td>
                        <% end if	%>
								<td><%=car_owner%></td>
								<td class="right"><%=formatnumber(overtime_cost,0)%></td>
								<td class="right"><%=formatnumber(general_cost,0)%></td>
								<td class="right"><%=formatnumber(fare_cost,0)%></td>
								<td class="right"><%=formatnumber(tot_km,0)%></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum_cost,0)%></td>
								<td class="right"><%=formatnumber(oil_cash_cost,0)%></td>
								<td class="right"><%=formatnumber(parking_cost,0)%></td>
								<td class="right"><%=formatnumber(toll_cost,0)%></td>
								<td class="right"><%=formatnumber(cash_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(return_cash,0)%></td>
								<td class="right"><%=formatnumber(repair_cost,0)%></td>
								<td class="right"><%=formatnumber(card_cost,0)%></td>
								<td class="right"><%=formatnumber(cost_sum,0)%></td>
							</tr>
						<%
							end if
							rs_emp.movenext()
						loop
						%>
							<tr>
								<td colspan="3" bgcolor="#EEFFFF" class="first">�Ұ�</td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_general_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_fare_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_km,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_somopum_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_oil_cash_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_parking_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_toll_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cash_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_juyoo_card_price,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_return_cash,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_repair_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_card_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cost_sum,0)%></td>
							</tr>
						<% if saupbu <> "����" then	%> 
							<tr>
								<td colspan="3" bgcolor="#FFE8E8" class="first">�Ѱ�</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_overtime_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_general_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_fare_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_tot_km,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_tot_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_somopum_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_oil_cash_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_parking_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_toll_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_cash_tot_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_juyoo_card_price,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_return_cash,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_repair_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_card_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_cost_sum,0)%></td>
							</tr>
						<% end if	%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
