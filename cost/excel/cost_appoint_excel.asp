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
Dim srchEmpMonth
Dim pre_year, pre_month, curr_year, curr_month, curr_date, pre_date
Dim title_line, savefilename, rsApp

srchEmpMonth = Request.QueryString("srchEmpMonth")

curr_year = Mid(srchEmpMonth, 1, 4)
curr_month = Mid(srchEmpMonth, 5, 6)

If (curr_month - 1) = 0 Then
	pre_year = curr_year - 1
	pre_month = 12
Else
	pre_year = curr_year
	pre_month = curr_month - 1
End If

curr_date = CStr(curr_year & "-" & curr_month) & "-15"
pre_date = CStr(pre_year & "-" & pre_month) & "-16"

title_line = srchEmpMonth & " - 인사 이동발령 현황"
savefilename = title_line & ".xls"

'엑셀 지정
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT app_empno, app_emp_name, app_date, app_to_company, "
objBuilder.Append "	app_to_org, app_to_orgcode, app_to_grade, app_to_position, app_be_company, "
objBuilder.Append "	app_be_org, app_be_orgcode, app_be_grade, app_be_position "
objBuilder.Append "FROM emp_appoint "
objBuilder.Append "WHERE (app_date >= '"&pre_date&"' AND app_date <= '"&curr_date&"') AND app_empno < '900000' "
objBuilder.Append "	AND app_id = '이동발령' "
objBuilder.Append "ORDER BY app_date, app_empno ASC "

Set rsApp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=title_line%></title>
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
		<td colspan="12" bgcolor="#FFFFFF">
			<div align="left" class="style2">&nbsp;<%=pre_date%>&nbsp;∼&nbsp;<%=curr_date%> &nbsp;인사 이동발령 현황&nbsp;</div>
		</td>
	</tr>
	<tr bgcolor="#EFEFEF" class="style11">
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성명</div></td>
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령일</div></td>
		<td colspan="3" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">발령전</div></td>
		<td colspan="3" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">발령후</div></td>
	</tr>
	<tr>
		<td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">회사</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">소속</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">직급/책</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">회사</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">소속</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">직급/책</div></td>
	</tr>
	<%
	  Do Until rsApp.EOF
	%>
	<tr valign="middle" class="style11">
		<td width="95"><div align="center" class="style1"><%=rsApp("app_empno")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rsApp("app_emp_name")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rsApp("app_date")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rsApp("app_to_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rsApp("app_to_org")%>(<%=rsApp("app_to_orgcode")%>)</div></td>
		<td width="145"><div align="center" class="style1"><%=rsApp("app_to_grade")%>-<%=rsApp("app_to_position")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rsApp("app_be_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rsApp("app_be_org")%>(<%=rsApp("app_be_orgcode")%>)</div></td>
		<td width="145"><div align="center" class="style1"><%=rsApp("app_be_grade")%>-<%=rsApp("app_be_position")%></div></td>
	</tr>
	<%
		rsApp.MoveNext()
	Loop

	rsApp.Close() : Set rsApp = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>