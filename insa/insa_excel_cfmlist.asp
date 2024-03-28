<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim from_date, to_date, company, cfm_type, com_sql, type_sql
Dim curr_date, cfm_company, savefilename, rsCfm

from_date = Request("from_date")
to_date = Request("to_date")
company = Request("company")
cfm_type = Request("cfm_type")

If company = "전체" Then
	com_sql = ""
Else
  	com_sql = "AND cfm_company ='" & company & "' "
End If

If cfm_type = "전체" Then
	type_sql = ""
Else
  	type_sql = "AND cfm_type ='" & cfm_type & "' "
End If

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = "제증명 발급현황 -- " & cfm_company & "" & cfm_type & "" & CStr(curr_date) & ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

objBuilder.Append "SELECT ecft.cfm_empno, ecft.cfm_emp_name, ecft.cfm_company, ecft.cfm_org_name, ecft.cfm_date, "
objBuilder.Append "	ecft.cfm_number, ecft.cfm_seq, ecft.cfm_type, ecft.cfm_use, ecft.cfm_use_dept, "
objBuilder.Append "	ecft.cfm_person1, ecft.cfm_person2, ecft.cfm_comment, ecft.cfm_job, ecft.cfm_position "
'objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
'objBuilder.Append "	emtt.emp_org_code, emp_reside_place, emp_reside_company, "
'objBuilder.Append "	emtt.emp_person1, emtt.emp_person2, "
'objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_team "
objBuilder.Append "FROM emp_confirm AS ecft "
objBuilder.Append "INNER JOIN emp_master AS emtt ON ecft.cfm_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (ecft.cfm_date >= '"&from_date&"' AND ecft.cfm_date <= '"&to_date&"') "
objBuilder.Append com_sql & type_sql
objBuilder.Append "ORDER BY cfm_type, cfm_seq DESC "

Set rsCfm = Server.CreateObject("ADODB.RecordSet")
rsCfm.Open objBuilder.ToString(), DBConn, 1
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;제증명 발급현황>&nbsp;(<%=cfm_company%>)&nbsp;<%=cfm_type%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">주민등록번호</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">제증명</div></td>
    <td><div align="center" class="style1">발급일자</div></td>
    <td><div align="center" class="style1">용도</div></td>
    <td><div align="center" class="style1">사용처</div></td>
    <td><div align="center" class="style1">비고</div></td>
  </tr>
    <%
	Do Until rsCfm.EOF
	%>
  <tr valign="middle" class="style11">
    <td width="59"><div align="center" class="style1"><%=rsCfm("cfm_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCfm("cfm_emp_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_person1")%>-<%=rsCfm("cfm_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsCfm("cfm_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsCfm("cfm_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_use")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCfm("cfm_use_dept")%></div></td>
    <td width="200"><div align="center" class="style1"><%=rsCfm("cfm_comment")%></div></td>
  </tr>
	<%
		rsCfm.MoveNext()
	Loop
	rsCfm.Close() : Set rsCfm = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>