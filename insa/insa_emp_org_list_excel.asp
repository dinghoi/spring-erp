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
Dim view_condi, curr_date, savefilename
Dim order_sql, where_sql, rs

view_condi = Request("view_condi")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = view_condi + " 조직별 인원현황" + CStr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정2021-03-02
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

order_Sql = "ORDER BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name, emtt.emp_no, emtt.emp_in_date "
where_sql = "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "

If view_condi = "전체" Then
	where_sql = where_sql & "AND emtt.emp_no < '900000' "
Else
	where_sql = where_sql & "AND eomt.org_company = '"&view_condi&"' AND emtt.emp_no < '900000' "
End If

objBuilder.Append "SELECT eomt.org_code, eomt.org_level, eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name, "
objBuilder.Append "	eomt.org_reside_place, eomt.org_reside_company, "
objBuilder.Append "	emtt.emp_no, emtt.emp_name, emtt.emp_job, emtt.emp_position, emtt.emp_in_date, "
objBuilder.Append "	emtt.emp_org_baldate, emtt.emp_birthday, emtt.emp_stay_name, emtt.emp_sex, "
objBuilder.Append "	emtt.emp_type, emtt.emp_grade, emtt.cost_center "
objBuilder.Append "FROM emp_org_mst AS eomt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON eomt.org_code = emtt.emp_org_code "
objBuilder.Append where_sql & order_sql

Set rs = Server.CreateObject("ADODB.Recordset")
Rs.Open objBuilder.ToString(), DBConn, 1
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=view_condi%> &nbsp;인원 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">성별</div></td>
    <td><div align="center" class="style1">직원구분</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">본부</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">상주회사</div></td>
    <td><div align="center" class="style1">비용구분</div></td>
    <td><div align="center" class="style1">실근무지</div></td>
    <td><div align="center" class="style1">입사일</div></td>
    <td><div align="center" class="style1">생년월일</div></td>
  </tr>
    <%
		Do Until rs.EOF
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_sex")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_type")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_reside_company")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("cost_center")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rs("emp_stay_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_birthday")%></div></td>
  </tr>
	<%
			Rs.MoveNext()
		Loop

		rs.Close() : Set rs = Nothing
		DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>