<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim owner_view, field_check, field_view, curr_date
Dim owner_gubun, savefilename
Dim owner_sql, rsCar, sqlWhereStr

owner_view = Request("owner_view")
field_check = Request("field_check")
field_view = Request("field_view")

curr_date = datevalue(mid(cstr(now()),1,10))

sqlWhereStr = "WHERE (end_date <> '' AND end_date <> '1900-01-01') "

Select Case owner_view
	Case "C"
		owner_gubun = "ȸ�� "
		owner_sql = "AND car_owner = 'ȸ��' "
	Case "P"
		owner_gubun = "���� "
		owner_sql = "AND car_owner = '����' "
	Case Else
		owner_gubun = "��ü"
		owner_sql = "AND (car_owner = '����' OR car_owner = 'ȸ��') "
End Select

savefilename = owner_gubun & " ���� ��Ȳ " & CStr(curr_date) & ".xls"

Call ViewExcelType(savefilename)

objBuilder.Append "SELECT car_no, car_name, car_year, oil_kind, car_company, car_use_dept, "
objBuilder.Append "	car_use, owner_emp_name, owner_emp_no, car_reg_date, last_km, "
objBuilder.Append "	insurance_date, insurance_company, insurance_amt, last_check_date, "
objBuilder.Append "	car_status, car_comment, end_date "
objBuilder.Append "FROM car_info "
objBuilder.Append sqlWhereStr & owner_sql

If field_check <> "total" Then
	objBuilder.Append "AND (" & field_check & " LIKE '%" & field_view & "%') "
End If

objBuilder.Append "ORDER BY car_no DESC "

Set rsCar = Server.CreateObject("ADODB.RecordSet")
rsCar.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=owner_gubun%> &nbsp;���� ��Ȳ&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">������ȣ</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">��������ȸ��</div></td>
    <td><div align="center" class="style1">���μ�</div></td>
    <td><div align="center" class="style1">�뵵</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">���������</div></td>
	<td><div align="center" class="style1">����ó����</div></td>
    <td><div align="center" class="style1">����Km</div></td>
    <td><div align="center" class="style1">����Ⱓ</div></td>
    <td><div align="center" class="style1">����ȸ��</div></td>
    <td><div align="center" class="style1">�����</div></td>
    <td><div align="center" class="style1">����������</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <%' �Ʒ��κ��� �ϴ� ���Ƴ���... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">�԰� ���γ��� </div> %>
    <%'</div></td> %>
  </tr>
    <%
	Do Until rsCar.EOF
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rsCar("car_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCar("car_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCar("car_year")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCar("oil_kind")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCar("car_company")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCar("car_use_dept")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCar("car_use")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCar("owner_emp_name")%>(<%=rsCar("owner_emp_no")%>)&nbsp;</div></td>
    <td width="145"><div align="center" class="style1"><%=rsCar("car_reg_date")%></div></td>
	<td width="145"><div align="center" class="style1"><%=rsCar("end_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=FormatNumber(rsCar("last_km"), 0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCar("insurance_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCar("insurance_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=FormatNumber(rsCar("insurance_amt"), 0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCar("last_check_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCar("car_status")%></div></td>
    <td width="200"><div align="center" class="style1"><%=rsCar("car_comment")%></div></td>
  </tr>
	<%
	rsCar.MoveNext()
	Loop
	rsCar.Close() : Set rsCar = Nothing
	%>
</table>
</body>
</html>
<!--#include virtual="/common/inc_footer.asp" -->