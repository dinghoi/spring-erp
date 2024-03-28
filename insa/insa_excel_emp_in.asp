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
Dim view_condi, from_date, to_date, curr_date, savefilename
Dim where_sql, rs_emp
Dim stay_name, stay_code, emp_person2, sex_id, emp_sex
Dim emp_birthday, emp_military_date1, emp_military_date2
Dim emp_marry_date, emp_grade_date, emp_end_date, emp_org_baldate
Dim emp_sawo_date, emp_email, emp_sawo_id, rs_stay

view_condi = Request.QueryString("view_condi")
from_date = Request.QueryString("from_date")
to_date = Request.QueryString("to_date")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

savefilename = "입사자 현황 -- " & to_date & "" & view_condi & "" & CStr(curr_date) & ".xls"
Call ViewExcelType(savefilename)

If view_condi <> "전체" Then
	where_sql = "	AND eomt.org_company='" & view_condi & "' "
Else
	where_sql = ""
End If

objBuilder.Append "SELECT emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_no, emtt.emp_name, "
objBuilder.Append "	emtt.emp_birthday, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_in_date, emtt.emp_last_edu, emtt.emp_disabled, "
objBuilder.Append "	emtt.emp_disab_grade, emtt.emp_reside_company, emtt.emp_stay_code, "
objBuilder.Append "	emtt.emp_reside_place, emtt.emp_first_date, emtt.emp_gunsok_date, emtt.emp_marry_date, "
objBuilder.Append "	emtt.emp_end_gisan, emtt.emp_yuncha_date, emtt.emp_jikmu, emtt.emp_family_zip, "
objBuilder.Append "	emtt.emp_family_sido, emtt.emp_family_gugun, emtt.emp_family_dong, emtt.emp_family_addr, "
objBuilder.Append "	emtt.emp_person1, emtt.emp_person2, emtt.emp_military_date1, emtt.emp_military_date2, "
objBuilder.Append "	emtt.emp_zipcode, emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, "
objBuilder.Append "	emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_emergency_tel, "
objBuilder.Append "	emtt.emp_hp_ddd, emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_sawo_id, "
objBuilder.Append "	emtt.emp_end_date, emtt.emp_sawo_date, emtt.emp_email, emtt.emp_type, "
objBuilder.Append "	emtt.emp_disabled, emtt.emp_disab_grade, emtt.emp_hobby, emtt.emp_faith, "
objBuilder.Append "	emtt.emp_military_id, emtt.emp_military_grade, emtt.emp_military_comm, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (emtt.emp_in_date >= '" & from_date & "' AND emtt.emp_in_date <= '" & to_date & "') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append where_sql
objBuilder.Append "ORDER BY emtt.emp_no, emtt.emp_name ASC "

Set rs_emp = DBConn.Execute(objBuilder.ToString())
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=from_date%>&nbsp;∼&nbsp;<%=to_date%> &nbsp;입사자 현황>&nbsp;(<%=view_condi%>)</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">성별</div></td>
    <td><div align="center" class="style1">직원구분</div></td>
    <td><div align="center" class="style1">주민번호</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">본부</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">상주처</div></td>
    <td><div align="center" class="style1">상주회사</div></td>
    <td><div align="center" class="style1">실근무지</div></td>
    <td><div align="center" class="style1">최초입사일</div></td>
    <td><div align="center" class="style1">입사일</div></td>
    <td><div align="center" class="style1">근속기산일</div></td>
    <td><div align="center" class="style1">퇴직기산일</div></td>
    <td><div align="center" class="style1">연차기산일</div></td>
    <td><div align="center" class="style1">생년월일</div></td>
    <td><div align="center" class="style1">직무</div></td>
    <td><div align="center" class="style1">본적주소</div></td>
    <td><div align="center" class="style1">현주소</div></td>
    <td><div align="center" class="style1">전화번호</div></td>
    <td><div align="center" class="style1">핸드폰</div></td>
    <td><div align="center" class="style1">e메일</div></td>
    <td><div align="center" class="style1">비상연락망</div></td>
    <td><div align="center" class="style1">경조회</div></td>
    <td><div align="center" class="style1">장애여부</div></td>
    <td><div align="center" class="style1">병역사항</div></td>
    <td><div align="center" class="style1">취미</div></td>
    <td><div align="center" class="style1">종교</div></td>
    <td><div align="center" class="style1">결혼기념일</div></td>
  </tr>
    <%
	Do Until rs_emp.EOF
		stay_name = ""
		stay_code = rs_emp("emp_stay_code")

        If stay_code <> "" Then
		   objBuilder.Append "SELECT stay_name FROM emp_stay WHERE stay_code = '"&stay_code&"'"
		   Set rs_stay = DBConn.Execute(objBuilder.ToString())

		  If Not rs_stay.EOF Then
             stay_name = rs_stay("stay_name")
		  End If
		  rs_stay.Close()
		End If

		emp_person2 = rs_emp("emp_person2")

        If emp_person2 <> "" Then
	       sex_id = Mid(CStr(emp_person2), 1, 1)

			If sex_id = "1" Then
	             emp_sex = "남"
			Else
	    	     emp_sex = "여"
			End If
	    End If

		If rs_emp("emp_birthday") = "1900-01-01" Then
			emp_birthday = ""
		Else
			emp_birthday = rs_emp("emp_birthday")
		End If

		If rs_emp("emp_military_date1") = "1900-01-01" Then
		   emp_military_date1 = ""
		   emp_military_date2 = ""
		Else
		   emp_military_date1 = rs_emp("emp_military_date1")
		   emp_military_date2 = rs_emp("emp_military_date2")
		End If

		If rs_emp("emp_marry_date") = "1900-01-01" Then
			emp_marry_date = ""
		Else
			emp_marry_date = rs_emp("emp_marry_date")
		End If

		If rs_emp("emp_grade_date") = "1900-01-01" Then
			emp_grade_date = ""
		Else
			emp_grade_date = rs_emp("emp_grade_date")
		End If

		If rs_emp("emp_end_date") = "1900-01-01" Then
			emp_end_date = ""
		Else
			emp_end_date = rs_emp("emp_end_date")
		End If

		If rs_emp("emp_org_baldate") = "1900-01-01" Then
			emp_org_baldate = ""
		Else
			emp_org_baldate = rs_emp("emp_org_baldate")
		End If

		If rs_emp("emp_sawo_date") = "1900-01-01" Then
			emp_sawo_date = ""
		Else
			emp_sawo_date = rs_emp("emp_sawo_date")
		End If

		emp_email = rs_emp("emp_email")&"@k-won.co.kr"
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_sex%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_person1")%>-<%=rs_emp("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs_emp("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs_emp("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs_emp("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("org_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("org_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("org_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_reside_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_reside_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=stay_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_gunsok_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_end_gisan")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_yuncha_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_birthday%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs_emp("emp_jikmu")%></div></td>
    <td width="350">
		<div align="center" class="style1">
			<%=rs_emp("emp_family_zip")%>&nbsp;<%=rs_emp("emp_family_sido")%>&nbsp;<%=rs_emp("emp_family_gugun")%>&nbsp;<%=rs_emp("emp_family_dong")%>&nbsp;<%=rs_emp("emp_family_addr")%>
		</div>
	</td>
    <td width="350">
		<div align="center" class="style1">
			<%=rs_emp("emp_zipcode")%>&nbsp;<%=rs_emp("emp_sido")%>&nbsp;<%=rs_emp("emp_gugun")%>&nbsp;<%=rs_emp("emp_dong")%>&nbsp;<%=rs_emp("emp_addr")%>
		</div>
	</td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_tel_ddd")%>-<%=rs_emp("emp_tel_no1")%>-<%=rs_emp("emp_tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_hp_ddd")%>-<%=rs_emp("emp_hp_no1")%>-<%=rs_emp("emp_hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_email%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_emergency_tel")%></div></td>
    <%
	If rs_emp("emp_sawo_id") = "Y" Then
	   emp_sawo_id = "가입"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%>-<%=emp_sawo_date%></div></td>
    <%
	Else
	   emp_sawo_id = "안함"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%></div></td>
    <%
	End If

	%>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_disabled")%>&nbsp;<%=rs_emp("emp_disab_grade")%></div></td>
    <td width="200">
		<div align="center" class="style1">
			<%=rs_emp("emp_military_id")%>&nbsp;<%=emp_military_date1%>&nbsp;<%=emp_military_date2%>&nbsp;<%=rs_emp("emp_military_grade")%>&nbsp;<%=rs_emp("emp_military_comm")%>
		</div>
	</td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_hobby")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs_emp("emp_faith")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_marry_date%></div></td>
  </tr>
	<%
		rs_emp.MoveNext()
	Loop
	Set rs_stay = Nothing
	rs_emp.Close() : Set rs_emp = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>