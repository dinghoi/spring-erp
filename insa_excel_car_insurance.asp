<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")
owner_view=Request("owner_view")
from_date=Request("from_date")
to_date=Request("to_date")

response.write(owner_view)
	
curr_date = datevalue(mid(cstr(now()),1,10))

if owner_view = "C" then
	title_line = cstr(from_date) + "~ " + cstr(to_date) + " " + " ���� ���踸�� ������Ȳ"
  else
	title_line = cstr(from_date) + "~ " + cstr(to_date) + " " + " ���� ���� ������Ȳ"
end if

savefilename = title_line +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if owner_view = "C" then
	owner_sql = " ins_last_date >= '"+from_date+"' and ins_last_date <= '"+to_date+"' "
  else
	owner_sql = " ins_date >= '"+from_date+"' and ins_date <= '"+to_date+"' "
end if

order_sql = " ORDER BY ins_car_no,ins_date DESC"

if view_condi = "��ü" then
      Sql = "select * from car_insurance where " + owner_sql + order_sql
   else  
      Sql = "select * from car_insurance where ins_car_no = '"+view_condi+"' and " + owner_sql + order_sql
end if
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">������ȣ</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">���������</div></td>
    <td><div align="center" class="style1">��������ȸ��</div></td>
    <td><div align="center" class="style1">���μ�</div></td>
    <td><div align="center" class="style1">�뵵</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">���谡����</div></td>
    <td><div align="center" class="style1">���踸����</div></td>
    <td><div align="center" class="style1">����ȸ��</div></td>
    <td><div align="center" class="style1">�����</div></td>
    <td><div align="center" class="style1">����1</div></td>
    <td><div align="center" class="style1">����2</div></td>
    <td><div align="center" class="style1">�빰</div></td>
    <td><div align="center" class="style1">�ڱ⺸��</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����⵿</div></td>
    <td><div align="center" class="style1">��೻��</div></td>
    <%' �Ʒ��κ��� �ϴ� ���Ƴ���... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">�԰� ���γ��� </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 
           car_no = rs("ins_car_no")
							  
		   Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
           Set rs_car = DbConn.Execute(SQL)
		   if not rs_car.eof then
				car_name = rs_car("car_name")
				car_year = rs_car("car_year")
				car_reg_date = rs_car("car_reg_date")
				car_use_dept = rs_car("car_use_dept")
				car_company = rs_car("car_company")
				car_use = rs_car("car_use")
				owner_emp_name = rs_car("owner_emp_name")
				owner_emp_no = rs_car("owner_emp_no")
				oil_kind = rs_car("oil_kind")
	          else
			    car_name = ""
				car_year = ""
				car_reg_date = ""
				car_use_dept = ""
				car_company = ""
				car_use = ""
				owner_emp_name = ""
				owner_emp_no = ""
				oil_kind = ""
           end if
           rs_car.close()
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("ins_car_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_year%></div></td>
    <td width="115"><div align="center" class="style1"><%=oil_kind%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_reg_date%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_company%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_use_dept%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_use%></div></td>
    <td width="115"><div align="center" class="style1"><%=owner_emp_name%>(<%=owner_emp_no%>)&nbsp;</div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ins_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("ins_last_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(rs("ins_amount"),0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("ins_man1")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("ins_man2")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("ins_object")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ins_self")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ins_injury")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ins_self_car")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ins_age")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ins_scramble")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
<% if rs("ins_contract_yn") = "Y" then %>
   <td width="145"><div align="center" class="style1">��೻������</div></td>
<%    else %>
   <td width="145"><div align="center" class="style1">��೻�������(<%=rs("ins_comment")%>)</div></td>
<% end if %>
  </tr>
	<%
	Rs.MoveNext()
	loop
	%>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
