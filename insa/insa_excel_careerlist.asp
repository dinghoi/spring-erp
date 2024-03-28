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
Dim view_condi, condi, curr_date, savefilename
Dim rsCareer
Dim career_empno, emp_name, emp_grade, emp_job
Dim emp_position, emp_org_code, emp_org_name
Dim emp_team, emp_reside_place, emp_bonbu
Dim emp_reside_company, emp_person1, emp_person2

view_condi = Request("view_condi")
condi = Request("condi")

If view_condi = "전체" Then
	condi = ""
End If

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = "경력현황 -- " & condi & "" & view_condi & "" & CStr(curr_date) & ".xls"

Call ViewExcelType(savefilename)

objBuilder.Append "SELECT emct.career_task, emct.career_empno, emct.career_office, emct.career_join_date, "
objBuilder.Append "	emct.career_end_date, emct.career_dept, emct.career_position, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_company, emtt.emp_reside_place, "
objBuilder.Append "	emtt.emp_reside_company, emtt.emp_person1, emtt.emp_person2, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_team "
objBuilder.Append "FROM emp_career AS emct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emct.career_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "

If view_condi = "상주처회사" Then
	objBuilder.Append "AND emtt.emp_reside_company  "
ElseIf view_condi = "경력업무" then
	objBuilder.Append "AND emct.career_task "
Else
	objBuilder.Append "AND emtt.emp_name "
End If

objBuilder.Append "LIKE '%"&condi&"%' "
objBuilder.Append "ORDER BY emct.career_empno ASC "

Set rsCareer = DBConn.Execute(objBuilder.ToString())
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;경력 현황>&nbsp;(<%=condi%>)&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">주민등록번호</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">회사</div></td>
	<td><div align="center" class="style1">본부</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">상주처회사</div></td>
    <td><div align="center" class="style1">경력회사</div></td>
    <td><div align="center" class="style1">경력기간</div></td>
    <td><div align="center" class="style1">부서</div></td>
    <td><div align="center" class="style1">직위/직책</div></td>
    <td><div align="center" class="style1">주요업무</div></td>
  </tr>
    <%
		Do Until rsCareer.EOF
			career_empno = rsCareer("career_empno")
			emp_name = rsCareer("emp_name")
			emp_grade = rsCareer("emp_grade")
			emp_job = rsCareer("emp_job")
			emp_position = rsCareer("emp_position")
			emp_org_code = rsCareer("emp_org_code")
			emp_org_name = rsCareer("org_name")
			emp_company = rsCareer("org_company")
			emp_bonbu = rsCareer("org_bonbu")
			emp_team = rsCareer("org_team")
			emp_reside_place = rsCareer("emp_reside_place")
			emp_reside_company = rsCareer("emp_reside_company")
			emp_person1 = rsCareer("emp_person1")
			emp_person2 = rsCareer("emp_person2")
	%>
  <tr valign="middle" class="style11">
    <td width="59"><div align="center" class="style1"><%=rsCareer("career_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_person1%>-<%=emp_person2%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_job%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_company%></div></td>
	<td width="145"><div align="center" class="style1"><%=emp_bonbu%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_team%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_org_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_reside_company%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCareer("career_office")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCareer("career_join_date")%>∼<%=rsCareer("career_end_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsCareer("career_dept")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsCareer("career_position")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rsCareer("career_task")%></div></td>
  </tr>
	<%
		rsCareer.MoveNext()
	Loop
	rsCareer.Close() : Set rsCareer = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>