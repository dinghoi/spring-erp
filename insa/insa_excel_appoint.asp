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
Dim view_condi, app_id, from_date, to_date, curr_date
Dim title_line, savefilename, rs

view_condi = Request("view_condi")
app_id = Request("app_id")
from_date = Request("from_date")
to_date = Request("to_date")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

title_line = view_condi & "(" & app_id & ") - 인사발령 현황(" & from_date & " ∼ " & to_date & ")"

savefilename = title_line & ".xls"

'엑셀 지정
Call ViewExcelType(savefilename)

'If view_condi = "전체" Then
	'If app_id = "전체" Then
		'Sql = "select * from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	'Else
		'Sql = "select * from emp_appoint where app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	'End If
'Else
	'If app_id = "전체" Then
		'Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	'Else
		'Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	'End If
'End If

objBuilder.Append "SELECT app_empno, app_emp_name, app_date, app_id, app_id_type, app_to_company, "
objBuilder.Append "	app_to_org, app_to_orgcode, app_to_grade, app_to_position, app_be_company, "
objBuilder.Append "	app_be_org, app_be_orgcode, app_be_grade, app_be_position, app_start_date, "
objBuilder.Append "	app_finish_date, app_be_enddate, app_reward, app_comment "
objBuilder.Append "FROM emp_appoint "
objBuilder.Append "WHERE (app_date >= '"&from_date&"' AND app_date <= '"&to_date&"') AND app_empno < '900000' "

If view_condi <> "전체" Then
	objBuilder.Append "	AND app_to_company = '"&view_condi&"' "
End If

If app_id <> "전체" Then
	objBuilder.Append "	AND app_id = '"&app_id&"' "
End If

objBuilder.Append "ORDER BY app_date, app_empno ASC "

'Set rs = Server.CreateObject("ADODB.RecordSet")
'Rs.Open Sql, Dbconn, 1
Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html lang="ko">
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
			<div align="left" class="style2">&nbsp;<%=from_date%>&nbsp;∼&nbsp;<%=to_date%> &nbsp;인사발령 현황>&nbsp;(<%=view_condi%>)</div>
		</td>
	</tr>
	<tr bgcolor="#EFEFEF" class="style11">
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성명</div></td>
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령일</div></td>
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령구분</div></td>
		<td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령유형</div></td>
		<td colspan="3" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">발령전</div></td>
		<td colspan="4" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">발령후</div></td>
	</tr>
	<tr>
		<td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">회사</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">소속</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">직급/책</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">회사</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">소속</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">직급/책</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">발령내용</div></td>
	</tr>
	<%
	  Do Until rs.EOF
	%>
	<tr valign="middle" class="style11">
		<td width="95"><div align="center" class="style1"><%=rs("app_empno")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rs("app_emp_name")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rs("app_date")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rs("app_id")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rs("app_id_type")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rs("app_to_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs("app_to_org")%>(<%=rs("app_to_orgcode")%>)</div></td>
		<td width="145"><div align="center" class="style1"><%=rs("app_to_grade")%>-<%=rs("app_to_position")%></div></td>
		<td width="95"><div align="center" class="style1"><%=rs("app_be_company")%></div></td>
		<td width="145"><div align="center" class="style1"><%=rs("app_be_org")%>(<%=rs("app_be_orgcode")%>)</div></td>
		<td width="145"><div align="center" class="style1"><%=rs("app_be_grade")%>-<%=rs("app_be_position")%></div></td>
		<td width="300" class="left">
			<div align="center" class="style1">
				<%=rs("app_start_date")%>&nbsp;-&nbsp;<%=rs("app_finish_date")%>&nbsp;<%=rs("app_be_enddate")%>&nbsp;<%=rs("app_reward")%>&nbsp;:&nbsp;<%=rs("app_comment")%>
			</div>
		</td>
	</tr>
	<%
		rs.MoveNext()
	Loop

	rs.Close() : Set rs = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>