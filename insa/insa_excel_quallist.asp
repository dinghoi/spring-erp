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
Dim view_company, view_condi, condi, curr_date, savefilename
Dim rsQual, rs_emp
Dim emp_name, emp_grade, emp_job, emp_position, emp_org_code
Dim emp_org_name, emp_team, emp_reside_place
Dim emp_reside_company, emp_person1, emp_person2
Dim qual_empno

view_company = Request("view_company")
view_condi = Request("view_condi")
condi = Request("condi")

If view_condi = "전체" Then
	condi = ""
End If

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = "자격증 보유현황 -- " & condi &""& view_condi &"" & cstr(curr_date) & ".xls"

Call ViewExcelType(savefilename)

objBuilder.Append "SELECT emqt.qual_empno, emqt.qual_type, emqt.qual_grade, emqt.qual_org, emqt.qual_no, emqt.qual_pass_date, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_company, emtt.emp_team, "
objBuilder.Append "	emtt.emp_reside_place, emtt.emp_reside_company, emtt.emp_person1, "
objBuilder.Append "	emtt.emp_person2, eomt.org_name, eomt.org_company, eomt.org_team, "
objBuilder.Append "	eomt.org_reside_place, eomt.org_reside_company "
objBuilder.Append "FROM emp_qual AS emqt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emqt.qual_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND eomt.org_company LIKE '%"&view_company&"%' "

If view_condi = "상주처회사" Then
	objBuilder.Append "AND eomt.org_reside_place "
ElseIf view_condi = "자격증명" Then
	objBuilder.Append "AND emqt.qual_type "
Else
	objBuilder.Append "AND emtt.emp_name "
End If

objBuilder.Append "LIKE '%"&condi&"%' "
objBuilder.Append "ORDER BY emqt.qual_empno ASC "

Set rsQual = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=NOW()%> &nbsp;자격증 보유현황>&nbsp;(<%=condi%>)&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">자격종목</div></td>
    <td><div align="center" class="style1">등급</div></td>
    <td><div align="center" class="style1">발급기관</div></td>
    <td><div align="center" class="style1">자격등록번호</div></td>
    <td><div align="center" class="style1">취득일</div></td>
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">주민등록번호</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">상주처회사</div></td>
  </tr>
<%
Do Until rsQual.eof
	qual_empno = rsQual("qual_empno")
	emp_name = rsQual("emp_name")
	emp_grade = rsQual("emp_grade")
	emp_job = rsQual("emp_job")
	emp_position = rsQual("emp_position")
	emp_org_code = rsQual("emp_org_code")
	emp_org_name = rsQual("org_name")
	emp_company = rsQual("org_company")
	emp_team = rsQual("org_team")
	emp_reside_place = rsQual("org_reside_place")
	emp_reside_company = rsQual("org_reside_company")
	emp_person1 = rsQual("emp_person1")
	emp_person2 = rsQual("emp_person2")
%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="left" class="style1"><%=rsQual("qual_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsQual("qual_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsQual("qual_org")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rsQual("qual_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsQual("qual_pass_date")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsQual("qual_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_person1%>-<%=emp_person2%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_job%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_company%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_team%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_org_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_reside_company%></div></td>
  </tr>
	<%
		rsQual.MoveNext()
	Loop
	rsQual.Close() : Set rsQual = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>