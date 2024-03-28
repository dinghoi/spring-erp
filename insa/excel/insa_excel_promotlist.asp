<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'On Error Resume Next

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
Dim to_date, in_grade, in_company
Dim curr_date, curr_year, curr_month, curr_day
Dim savefilename, condi_sql, target_date, rs_emp

to_date = Request("to_date")
in_grade = Request("in_grade")
in_company = Request("in_company")

curr_date = Mid(CStr(Now()), 1, 10)
curr_year = Mid(CStr(Now()), 1, 4)
curr_month = Mid(CStr(Now()), 6, 2)
curr_day = Mid(CStr(Now()), 9, 2)

savefilename = "승진대상자 현황 -- " & in_company & "" & in_grade & "" & CStr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

If in_company = "" Then
	'in_company = "케이원정보통신"
	in_company = "전체"
	to_date = curr_date
	in_grade = "대리2급"
End If

Select Case in_grade
	Case "대리2급"
		condi_sql = "AND emp_grade LIKE '%사원%' "
	Case "대리1급"
		condi_sql = "AND emp_grade LIKE '%대리2급%' "
	Case "과장"
		condi_sql = "AND emp_grade LIKE '%대리2급%' OR emp_grade LIKE '%대리1급%' "
	Case "차장"
		condi_sql = "AND emp_grade LIKE '%과장%' "
	Case "부장"
		condi_sql = "AND emp_grade LIKE '%차장%' "
End Select

target_date = to_date

objBuilder.Append "SELECT emtt.emp_grade_date, emtt.emp_first_date, emtt.emp_no, emtt.emp_name, "
objBuilder.Append "	emtt.emp_birthday, emtt.emp_grade, emtt.emp_position, "
objBuilder.Append "	emtt.emp_in_date, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_team "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (ISNULL(emp_end_date) OR emp_end_date = '1900-01-01') "
objBuilder.Append "	AND emtt.emp_no < '999990' "
If in_company <> "전체" Then
	objBuilder.Append "	AND eomt.org_company = '"&in_company&"' "
End If
objBuilder.Append condi_sql
objBuilder.Append "ORDER BY emtt.emp_first_date, emtt.emp_no DESC "

Set rs_emp = Server.CreateObject("ADODB.RecordSet")
rs_emp.Open objBuilder.ToString(), DBConn, 1
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
		<td colspan="13" bgcolor="#FFFFFF">
			<div align="left" class="style2">&nbsp;<%=Now()%> &nbsp;승진대상자 현황>&nbsp;(<%=in_company%>)&nbsp;<%=in_grade%></div>
		</td>
	</tr>
	<tr bgcolor="#EFEFEF" class="style11">
		<td><div align="center" class="style1">사번</div></td>
		<td><div align="center" class="style1">성명</div></td>
		<td><div align="center" class="style1">생년월일</div></td>
		<td><div align="center" class="style1">현직급</div></td>
		<td><div align="center" class="style1">직책</div></td>
		<td><div align="center" class="style1">소속</div></td>
		<td><div align="center" class="style1">최초입사일</div></td>
		<td><div align="center" class="style1">입사일</div></td>
		<td><div align="center" class="style1">최종승진일</div></td>
		<td><div align="center" class="style1">대상년한</div></td>
		<td><div align="center" class="style1">회사</div></td>
		<td><div align="center" class="style1">본부</div></td>
		<td><div align="center" class="style1">팀</div></td>
	</tr>
	<%
	Dim emp_grade_date, year_cnt, mon_cnt, day_cnt, target_cnt

	Do Until rs_emp.EOF
		If rs_emp("emp_grade_date") = "1900-01-01" Then
		   emp_grade_date = ""
		Else
		   emp_grade_date = rs_emp("emp_grade_date")
		End If

		If emp_grade_date <> "" Then
			year_cnt = DateDiff("yyyy", rs_emp("emp_grade_date"), target_date)
			mon_cnt = DateDiff("m", rs_emp("emp_grade_date"), target_date)
			day_cnt = DateDiff("d", rs_emp("emp_grade_date"), target_date)
		Else
		   year_cnt = DateDiff("yyyy", rs_emp("emp_first_date"), target_date)
		   mon_cnt = DateDiff("m", rs_emp("emp_first_date"), target_date)
		   day_cnt = DateDiff("d", rs_emp("emp_first_date"), target_date)
		End If

		target_cnt = CInt(mon_cnt)

		If (in_grade = "대리2급" Or in_grade = "대리1급") And target_cnt > 24 Then
	%>
	<tr valign="middle" class="style11">
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_no")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("emp_name")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_birthday")%></div></td>
		<td width="59"><div align="center" class="style1"><%=rs_emp("emp_grade")%></div></td>
		<td width="59"><div align="center" class="style1"><%=rs_emp("emp_position")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_name")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_first_date")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_in_date")%></div></td>
		<td width="145"><div align="center" class="style1"><%=emp_grade_date%></div></td>
		<td width="145"><div align="center" class="style1"><%=mon_cnt%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_bonbu")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_team")%></div></td>
	</tr>
	<%
		Else
			If in_grade = "과장" And rs_emp("emp_grade") = "대리1급" And target_cnt > 36 Then
	%>
	<tr valign="middle" class="style11">
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_no")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("emp_name")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_birthday")%></div></td>
		<td width="59"><div align="center" class="style1"><%=rs_emp("emp_grade")%></div></td>
		<td width="59"><div align="center" class="style1"><%=rs_emp("emp_position")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_name")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_first_date")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_in_date")%></div></td>
		<td width="145"><div align="center" class="style1"><%=emp_grade_date%></div></td>
		<td width="145"><div align="center" class="style1"><%=mon_cnt%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_bonbu")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_team")%></div></td>
	</tr>
	<%
			Else
				If in_grade = "과장" And rs_emp("emp_grade") = "대리2급" And target_cnt > 36 Then
	%>
	<tr valign="middle" class="style11">
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_no")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("emp_name")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_birthday")%></div></td>
		<td width="59"><div align="center" class="style1"><%=rs_emp("emp_grade")%></div></td>
		<td width="59"><div align="center" class="style1"><%=rs_emp("emp_position")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_name")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_first_date")%></div></td>
		<td width="115"><div align="center" class="style1"><%=rs_emp("emp_in_date")%></div></td>
		<td width="145"><div align="center" class="style1"><%=emp_grade_date%></div></td>
		<td width="145"><div align="center" class="style1"><%=mon_cnt%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_bonbu")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs_emp("org_team")%></div></td>
	</tr>
	<%
	              End If
	        End If
		End If

		rs_emp.MoveNext()
	Loop
	rs_emp.Close() : Set rs_emp = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>