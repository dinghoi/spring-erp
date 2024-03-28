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
Dim view_condi, view_c, field_check, field_bonbu, field_team
Dim curr_date, savefilename
Dim rsOrg

view_condi = Request.QueryString("view_condi")
view_c = Request.QueryString("view_c")
field_check = Request.QueryString("field_check")
field_bonbu = Request.QueryString("field_bonbu")
field_team = Request.QueryString("field_team")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = view_condi&"조직현황"&CStr(curr_date)&".xls"

Call ViewExcelType(savefilename)

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "bonbu"
End If

objBuilder.Append "SELECT org_code, org_name, org_empno, org_emp_name, org_date, org_owner_empno, "
objBuilder.Append "org_owner_empname, org_company, org_bonbu, org_team, org_reside_place, org_reside_company, "
objBuilder.Append "org_level, org_cost_center "
objBuilder.Append "FROM emp_org_mst "

If field_check = "total" Then
	objBuilder.Append "WHERE (ISNULL(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '000-00-00') "
	objBuilder.Append "AND org_company = '"&view_condi&"' "

	field_check = ""
Else
	If view_c = "bonbu" Then
		objBuilder.Append "WHERE (ISNULL(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '000-00-00') "
		objBuilder.Append "AND org_company = '"&view_condi&"' "
		objBuilder.Append "AND org_bonbu LIKE '%"&field_bonbu&"%' "
	End If

	If view_c = "team" Then
		objBuilder.Append "WHERE (isNull(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '000-00-00') "
		objBuilder.Append "AND org_company = '"&view_condi&"' "
		objBuilder.Append "AND org_team LIKE '%"&field_team&"%' "
	End If

	If Trim(view_condi) = "케이네트웍스" Then
		objBuilder.Append "AND org_code > '6513' "
	End If
End If

objBuilder.Append "ORDER BY FIELD(org_level, '회사', '본부', '팀', '파트', '상주처') "

Set rsOrg = DBConn.Execute(objBuilder.ToString())
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=view_condi%> &nbsp;조직 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
	<td><div align="center" class="style1">조직코드</div></td>
    <td><div align="center" class="style1">조직명</div></td>
	<td><div align="center" class="style1">조직Lvel</div></td>
    <td><div align="center" class="style1">조직장사번</div></td>
    <td><div align="center" class="style1">조직장성명</div></td>
    <td><div align="center" class="style1">조직생성일</div></td>
    <td><div align="center" class="style1">상위조직장사번</div></td>
    <td><div align="center" class="style1">상위조직장성명</div></td>
    <td><div align="center" class="style1">소속회사</div></td>
    <td><div align="center" class="style1">소속본부</div></td>
    <td><div align="center" class="style1">소속팀</div></td>
    <td><div align="center" class="style1">상주처</div></td>
    <td><div align="center" class="style1">상주처회사</div></td>
    <td><div align="center" class="style1">비용구분</div></td>
  </tr>
    <%
	Do Until rsOrg.EOF
	%>
  <tr valign="middle" class="style11">
    <td width="59"><div align="center" class="style1"><%=rsOrg("org_code")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_name")%></div></td>
	<td width="145"><div align="center" class="style1"><%=rsOrg("org_level")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsOrg("org_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsOrg("org_emp_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsOrg("org_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsOrg("org_owner_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsOrg("org_owner_empname")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_reside_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_reside_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsOrg("org_cost_center")%></div></td>
  </tr>
	<%
		rsOrg.MoveNext()
	Loop
	rsOrg.Close() : Set rsOrg = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>