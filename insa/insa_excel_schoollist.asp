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
Dim condi_sql, rsSch

view_condi = Request("view_condi")
condi = Request("condi")

If view_condi = "전체" Then
	condi = ""
End If

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = "학력 현황 -- "&condi&view_condi&CStr(curr_date)&".xls"

Call ViewExcelType(savefilename)

Select Case view_condi
	Case "전체"
		condi_sql = ""
	Case "상주처회사"
		condi_sql = "AND emtt.emp_reside_company LIKE '%" & condi & "%' "
	Case "성명"
		condi_sql = "AND emtt.emp_name LIKE '%" & condi & "%' "
	Case Else
		condi_sql = "AND emct." & view_condi & " LIKE '%" & condi & "%' "
End Select

objBuilder.Append "SELECT emct.sch_empno, emct.sch_school_name, emct.sch_start_date, emct.sch_end_date, "
objBuilder.Append "	emct.sch_dept, emct.sch_major, emct.sch_sub_major, emct.sch_degree, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_org_code, "
objBuilder.Append "	emtt.emp_person1, emtt.emp_person2, emtt.emp_reside_company, emtt.emp_reside_place, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_team "
objBuilder.Append "FROM emp_school AS emct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emct.sch_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01') "
objBuilder.Append condi_sql
objBuilder.Append "ORDER BY emct.sch_empno ASC "

Set rsSch = DBConn.Execute(objBuilder.ToString())
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
    <td colspan="14" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;학력 현황>&nbsp;(<%=condi%>)&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">주민등록번호</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">상주처회사</div></td>
    <td><div align="center" class="style1">학교명</div></td>
    <td><div align="center" class="style1">기간</div></td>
    <td><div align="center" class="style1">학과</div></td>
    <td><div align="center" class="style1">전공</div></td>
    <td><div align="center" class="style1">부전공</div></td>
    <td><div align="center" class="style1">학위</div></td>
  </tr>
    <%
	Dim sch_empno, emp_name, emp_grade, emp_job, emp_position, emp_org_code
	Dim emp_org_name, emp_bonbu, emp_team, emp_reside_place
	Dim emp_reside_company, emp_person1, emp_person2
	Do Until rsSch.EOF
        sch_empno = rsSch("sch_empno")

		emp_name = rsSch("emp_name")
		emp_grade = rsSch("emp_grade")
		emp_job = rsSch("emp_job")
		emp_position = rsSch("emp_position")
		emp_org_code = rsSch("emp_org_code")
		emp_org_name = rsSch("org_name")
		emp_company = rsSch("org_company")
		emp_bonbu = rsSch("org_bonbu")
		emp_team = rsSch("org_team")
		emp_reside_place = rsSch("emp_reside_place")
		emp_reside_company = rsSch("emp_reside_company")
		emp_person1 = rsSch("emp_person1")
		emp_person2 = rsSch("emp_person2")
	%>
  <tr valign="middle" class="style11">
    <td width="59"><div align="center" class="style1"><%=rsSch("sch_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_person1%>-<%=emp_person2%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_job%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_company%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_team%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_org_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_reside_company%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsSch("sch_school_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsSch("sch_start_date")%>∼<%=rsSch("sch_end_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsSch("sch_dept")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsSch("sch_major")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsSch("sch_sub_major")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsSch("sch_degree")%></div></td>
  </tr>
	<%
		rsSch.MoveNext()
	Loop
	rsSch.Close() : Set rsSch = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>