<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "����� ��Ȳ" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY emp_company, emp_in_date,emp_no DESC"
'where_sql = " WHERE emp_disab_grade = 'Y' and isNull(emp_end_date) or emp_end_date = '1900-01-01'"
where_sql = " WHERE emp_disab_grade = 'Y'"

sql = "select * from emp_master " + where_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;����� ��Ȳ</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">���</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">�ֹι�ȣ</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��å</div></td>
    <td><div align="center" class="style1">�Ҽ�</div></td>
    <td><div align="center" class="style1">ȸ��</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">�����</div></td>
    <td><div align="center" class="style1">��</div></td>
    <td><div align="center" class="style1">����ó</div></td>
    <td><div align="center" class="style1">�Ի���</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">�������</div></td>
    <td><div align="center" class="style1">��ֿ���</div></td>
    <td><div align="center" class="style1">��ֵ��</div></td>
  </tr>
    <%
		do until rs.eof 
		
		if rs("emp_end_date") = "1900-01-01" then
		   emp_end_date = ""
		   else 
		   emp_end_date = rs("emp_end_date")
		end if
		if rs("emp_org_baldate") = "1900-01-01" then
		   emp_org_baldate = ""
		   else 
		   emp_org_baldate = rs("emp_org_baldate")
		end if
		if rs("emp_grade_date") = "1900-01-01" then
		   emp_grade_date = ""
		   else 
		   emp_grade_date = rs("emp_grade_date")
		end if

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_sex")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_person1")%>-<%=rs("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_reside_place")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_end_date%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_birthday")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_disabled")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_disab_grade")%></div></td>
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
