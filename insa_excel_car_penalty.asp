<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

field_check=Request("field_check")
field_view=Request("field_view")
from_date=Request("from_date")
to_date=Request("to_date")
	
curr_date = datevalue(mid(cstr(now()),1,10))

title_line = cstr(from_date) + "~ " + cstr(to_date) + " " + " ���� ���·� ��Ȳ"

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

If field_check = "total" Then
	field_view = ""
End If

owner_sql = " where pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' "
order_sql = " ORDER BY pe_car_no,pe_date,pe_seq DESC"

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

sql = "select * from car_penalty " + owner_sql + field_sql + order_sql
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
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">�μ�</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">���ݳ���</div></td>
    <td><div align="center" class="style1">���·�</div></td>
    <td><div align="center" class="style1">�������</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">���Աݾ�</div></td>
    <td><div align="center" class="style1">�뺸����</div></td>
    <td><div align="center" class="style1">�뺸���</div></td>
    <td><div align="center" class="style1">�̳�</div></td>
    <td><div align="center" class="style1">���</div></td>
    <%' �Ʒ��κ��� �ϴ� ���Ƴ���... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">�԰� ���γ��� </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 
           car_no = rs("pe_car_no")
		   
           if rs("pe_in_date") = "1900-01-01"  then
	               pe_in_date = ""
			  else 
			       pe_in_date = rs("pe_in_date")
	       end if
	       if rs("pe_notice_date") = "1900-01-01" then
	               pe_notice_date = ""
			  else 
			       pe_notice_date = rs("pe_notice_date")
	       end if
		   					  
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
    <td width="115"><div align="center" class="style1"><%=rs("pe_car_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=owner_emp_name%>(<%=owner_emp_no%>)</div></td>
    <td width="115"><div align="center" class="style1"><%=car_use_dept%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("pe_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("pe_comment")%></div></td>
    <td width="115"><div align="rigrh" class="style1"><%=formatnumber(rs("pe_amount"),0)%></div></td>
    <td width="200"><div align="center" class="style1"><%=rs("pe_place")%></div></td>
    <td width="115"><div align="center" class="style1"><%=pe_in_date%></div></td>
    <td width="115"><div align="rigrh" class="style1"><%=formatnumber(rs("pe_in_amt"),0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=pe_notice_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("pe_notice")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("pe_default")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("pe_bigo")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
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
