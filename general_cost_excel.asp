<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

slip_month = Request("slip_month")
view_c = Request("view_c")
view_d = Request("view_d")
emp_name = Request("emp_name")

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
    view_c = "total"
    view_d = "slip"
	emp_name = ""
end If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = slip_month

savefilename = slip_month + "�� �Ϲݰ�� ��Ȳ.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' �����Ǻ�
posi_sql = " and general_cost.emp_no = '" + user_id + "'"&chr(13)

if position = "����" then
	view_condi = "����"
end if

if position = "��Ʈ��" then
	if view_c = "total" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (org_name = '��ȭ����ȣ��' or org_name = '��ȭ��������') "&chr(13)
		  else
			posi_sql = " and org_name = '"&org_name&"'"&chr(13)
		end if
	  else
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (org_name = '��ȭ����ȣ��' or org_name = '��ȭ��������') and general_cost.emp_name like '%"&emp_name&"%'"&chr(13)
		  else
			posi_sql = " and org_name = '"&org_name&"' and general_cost.emp_name like '%"&emp_name&"%'"&chr(13)
		end if
	end if
end if

if position = "����" then
	if view_c = "total" then
        'posi_sql = " and team = '"&team&"'"
        posi_sql = " and (team = '"&team&"' or reside_place = '"&team&"') "&chr(13)
    else
        'posi_sql = " and team = '"&team&"' and general_cost.emp_name like '%"&emp_name&"%'"
        posi_sql = " and (team = '"&team&"' or reside_place = '"&team&"') and general_cost.emp_name like '%"&emp_name&"%' "&chr(13)
	end if
end if

if position = "�������" or cost_grade = "2" then
	if view_c = "total" then
        'posi_sql = " and saupbu = '"&saupbu&"'"
        posi_sql = " and saupbu = emp_master.emp_saupbu "&chr(13)
	else
        'posi_sql = " and saupbu = '"&saupbu&"' and emp_name like '%"&emp_name&"%'"
        posi_sql = " and saupbu = emp_master.emp_saupbu and general_cost.emp_name like '%" & emp_name & "%' "&chr(13)
	end if
end if

if position = "������" or cost_grade = "1" then
  	if view_c = "total" then
		posi_sql = " and bonbu = '"&bonbu&"'"&chr(13)
 	else
		posi_sql = " and bonbu = '"&bonbu&"' and general_cost.emp_name like '%"&emp_name&"%'"&chr(13)
	end if	 
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "��ü"
  	if view_c = "total" then
		posi_sql = " "
 	else
		posi_sql = " and general_cost.emp_name like '%"&emp_name&"%'"&chr(13)
	end if	 
end if

' ���Ǻ� ��ȸ.........
base_sql = "     select *                                           "&chr(13)&_
           "       from general_cost                                "&chr(13)&_
           " inner join emp_master                                  "&chr(13)&_           
           "         ON emp_master.emp_no =  general_cost.emp_no    "&chr(13)&_ 
           "      where (cost_reg = '0')                            "&chr(13)&_
           "        and (tax_bill_yn <> 'Y' or isnull(tax_bill_yn)) "&chr(13)&_
           "        and (slip_gubun = '���')                       "&chr(13)

if view_d = "slip" then
    base_sql  = base_sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
    order_sql = "  ORDER BY general_cost.org_name, general_cost.emp_name, general_cost.slip_date ASC "
end if
if view_d = "reg" then
    base_sql  = base_sql & " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
    order_sql = "  ORDER BY general_cost.org_name, general_cost.emp_name, general_cost.slip_date ASC "
end if

sql = base_sql + posi_sql + order_sql
Response.write "<pre>"&sql & "</pre><br>"

Rs.Open Sql, Dbconn, 1

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��</th>
								<th scope="col">������</th>
								<th scope="col">����ó</th>
								<th scope="col">���ȸ��</th>
								<th scope="col">���</th>
								<th scope="col">�����</th>
								<th scope="col">�߻�����</th>
								<th scope="col">�߱�����</th>
								<th scope="col">�������</th>
								<th scope="col">��뱸��</th>
								<th scope="col">����׸�</th>
								<th scope="col">����NO</th>
								<th scope="col">��û�ݾ�</th>
								<th scope="col">���ó</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							cost_sum = cost_sum + rs("cost")
							if rs("cancel_yn") = "Y" then
								cancel_sum = cancel_sum + rs("cost")
							end if
							if rs("cancel_yn") <> "Y" then
								if rs("pay_yn") = "Y" then
									pay_sum = pay_sum + rs("cost")
								  else
									mi_pay_sum = mi_pay_sum + rs("cost")
								end if
							end if							  

							if rs("pay_yn") = "Y" then
								pay_yn = "����"
							  else
							  	pay_yn = "������"
							end if
							if rs("cancel_yn") = "Y" then
								cancel_yn = "���"
							  else
							  	cancel_yn = "����"
							end if
						    %>
							<%
                            ' 5�� ���� ���� �Է°� ����...
                            chk_slip_month = mid(rs("slip_date"),1,7)
                            chk_reg_month = mid(rs("reg_date"),1,7)
                            chk_reg_day = mid(rs("reg_date"),9,2)

                            if ((chk_slip_month < chk_reg_month) and (chk_reg_day > "05")) then
                                bgcolor = "burlywood"
                            else
                                bgcolor = "#f8f8f8"
                            end if
                            %>
                            <tr style="background-color: <%=bgcolor%>;">
								<td class="first"><%=rs("emp_company")%></td>
								<td><%=rs("bonbu")%></td>
								<td><%=rs("saupbu")%></td>
								<td><%=rs("team")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("reside_place")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("emp_no")%></td>
								<td><%=rs("emp_name")%>&nbsp;<%=rs("emp_grade")%></td>
                                <td><%=rs("slip_date")%></td>
                                <td><%=mid(rs("reg_date"),1,10)%></td>
								<td><%=rs("cost_center")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("sign_no")%>&nbsp;</td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td><%=rs("customer")%></td>
								<td><%=pay_yn%></td>
								<td><%=cancel_yn%></td>
								<td><%=rs("pl_yn")%></td>
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

