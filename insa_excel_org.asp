<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs

view_condi=Request("view_condi")
view_c=Request("view_c")
field_check=Request("field_check")
field_bonbu=Request("field_bonbu")
field_saupbu=Request("field_saupbu")
field_team=Request("field_team")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = view_condi + "������Ȳ" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_reside_place ASC"

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "bonbu"
End If

If field_check = "total" Then
       owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"')"
	   field_check = ""
   else
       if view_c = "bonbu" Then
              owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"') and (org_bonbu like '%" + field_bonbu + "%')"
       end if
	   if view_c = "saupbu" Then
              owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"') and (org_saupbu like '%" + field_saupbu + "%')"
       end if
	   if view_c = "team" Then
              owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"') and (org_team like '%" + field_team + "%')"
       end if
End If

sql = "select * from emp_org_mst " + owner_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=view_condi%> &nbsp;���� ��Ȳ&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">�����ڵ�</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">����T.O</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">�����强��</div></td>
    <td><div align="center" class="style1">����������</div></td>
    <td><div align="center" class="style1">������������</div></td>
    <td><div align="center" class="style1">���������强��</div></td>
    <td><div align="center" class="style1">�Ҽ�ȸ��</div></td>
    <td><div align="center" class="style1">�ҼӺ���</div></td>
    <td><div align="center" class="style1">�Ҽӻ����</div></td>
    <td><div align="center" class="style1">�Ҽ���</div></td>
    <td><div align="center" class="style1">����ó</div></td>
    <td><div align="center" class="style1">����óȸ��</div></td>
    <td><div align="center" class="style1">��뱸��</div></td>
    <%' �Ʒ��κ��� �ϴ� ���Ƴ���... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">�԰� ���γ��� </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof
	%>
  <tr valign="middle" class="style11">
    <td width="59"><div align="center" class="style1"><%=rs("org_code")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("org_table_org")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("org_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("org_emp_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("org_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("org_owner_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("org_owner_empname")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_reside_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_reside_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_cost_center")%></div></td>
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
