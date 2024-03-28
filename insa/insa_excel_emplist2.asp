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
Dim view_condi, curr_date, savefilename, condi_name
Dim rsEmp, emp_person2, sex_id, emp_sex, emp_birthday
Dim emp_military_date, emp_military_date1, emp_military_date2
Dim emp_marry_date, emp_grade_date, emp_end_date, emp_org_baldate
Dim emp_sawo_date, emp_email, emp_sawo_id
Dim stay_name, stay_code, rsStay

view_condi = Request.QueryString("view_condi")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

If view_condi = "" Then
	view_condi = "emp_image"
End If

Select Case view_condi
	Case "cost_center"
		condi_name = "비용배분구분"
	Case "emp_ename"
		condi_name = "성명(영문)"
	Case "emp_person1"
		condi_name = "주민등록번호"
	Case "emp_birthday"
		condi_name = "생년월일"
	Case "emp_sido"
		condi_name = "주소"
	Case "emp_tel_no1"
		condi_name = "전화번호"
	Case "emp_hp_no1"
		condi_name = "휴대폰번호"
	Case "emp_emergency_tel"
		condi_name = "비상연락"
	Case "emp_email"
		condi_name = "이메일"
	Case "emp_extension_no"
		condi_name = "내선번호"
	Case "emp_last_edu"
		condi_name = "최종학력"
	Case Else
		condi_name = "사진"
End Select

savefilename = "자료 미등록현황 -- "&condi_name&CStr(curr_date)&".xls"
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT emtt.emp_stay_code, emtt.emp_person2, emtt.emp_birthday, emtt.emp_military_date1, emtt.emp_marry_date, "
objBuilder.Append "	emtt.emp_grade_date, emtt.emp_end_date, emtt.emp_org_baldate, emtt.emp_sawo_date, emtt.emp_email, "
objBuilder.Append "	emtt.emp_no, emtt.emp_name, emtt.emp_type, emtt.emp_person1, emtt.emp_person2, emtt.emp_grade, "
objBuilder.Append "	emtt.emp_job, emtt.emp_position, emtt.emp_org_name, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, "
objBuilder.Append "	emtt.emp_team, emtt.emp_reside_place, emtt.emp_first_date, emtt.emp_in_date, emtt.emp_gunsok_date, "
objBuilder.Append "	emtt.emp_end_gisan, emtt.emp_yuncha_date, emtt.emp_jikmu, emtt.emp_last_edu, emtt.emp_family_zip, "
objBuilder.Append "	emtt.emp_family_sido, emtt.emp_family_gugun, emtt.emp_family_dong, emtt.emp_family_addr, emtt.emp_zipcode, "
objBuilder.Append "	emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, emtt.emp_tel_ddd, emtt.emp_tel_no1, "
objBuilder.Append "	emtt.emp_tel_no2, emtt.emp_hp_ddd, emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_emergency_tel, "
objBuilder.Append "	emtt.emp_sawo_id, emtt.emp_disabled, emtt.emp_disab_grade, emtt.emp_military_id, emtt.emp_military_grade, "
objBuilder.Append "	emtt.emp_military_date2, emtt.emp_military_comm, emtt.emp_hobby, emtt.emp_faith, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_reside_place "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append " AND (emtt." & view_condi & " = '' OR isNull(emtt." & view_condi & ")) "
objBuilder.Append "ORDER BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_code, emtt.emp_in_date, emtt.emp_no ASC "

Set rsEmp = DBConn.Execute(objBuilder.ToString())
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=Now()%> &nbsp;자료 미등록현황>&nbsp;(<%=condi_name%>)</div></td>
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
    <td><div align="center" class="style1">사업부</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">상주처</div></td>
    <td><div align="center" class="style1">실근무지</div></td>
    <td><div align="center" class="style1">최초입사일</div></td>
    <td><div align="center" class="style1">입사일</div></td>
    <td><div align="center" class="style1">근속기산일</div></td>
    <td><div align="center" class="style1">퇴직기산일</div></td>
    <td><div align="center" class="style1">연차기산일</div></td>
    <td><div align="center" class="style1">소속발령일</div></td>
    <td><div align="center" class="style1">승진일</div></td>
    <td><div align="center" class="style1">생년월일</div></td>
    <td><div align="center" class="style1">직무</div></td>
    <td><div align="center" class="style1">최종학력</div></td>
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
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
<%
Do Until rsEmp.EOF
	stay_name = ""
	stay_code = rsEmp("emp_stay_code")

	if stay_code <> "" then
	   objBuilder.Append "SELECT stay_name FROM emp_stay WHERE stay_code = '"&stay_code&"'"

	   Set rsStay = DBConn.Execute(objBuilder.ToString())
	   objBuilder.Clear()

	  If Not rsStay.eof Then
		 stay_name = rsStay("stay_name")
	  End If

	  rsStay.Close()
	End If

	emp_person2 = rsEmp("emp_person2")

	If emp_person2 <> "" Then
	   sex_id = Mid(CStr(emp_person2), 1, 1)

		If sex_id = "1" Then
			 emp_sex = "남"
		Else
			 emp_sex = "여"
		End If
	End If

	If rsEmp("emp_birthday") = "1900-01-01" Then
		emp_birthday = ""
	Else
		emp_birthday = rsEmp("emp_birthday")
	End If

	If rsEmp("emp_military_date1") = "1900-01-01" Then
		emp_military_date1 = ""
		emp_military_date2 = ""
	Else
		emp_military_date1 = rsEmp("emp_military_date1")
		emp_military_date2 = rsEmp("emp_military_date2")
	End If

	If rsEmp("emp_marry_date") = "1900-01-01" Then
		emp_marry_date = ""
	Else
		emp_marry_date = rsEmp("emp_marry_date")
	End If

	If rsEmp("emp_grade_date") = "1900-01-01" Then
		emp_grade_date = ""
	Else
		emp_grade_date = rsEmp("emp_grade_date")
	End If

	If rsEmp("emp_end_date") = "1900-01-01" Then
		emp_end_date = ""
	Else
		emp_end_date = rsEmp("emp_end_date")
	End If

	If rsEmp("emp_org_baldate") = "1900-01-01" Then
		emp_org_baldate = ""
	Else
		emp_org_baldate = rsEmp("emp_org_baldate")
	End If

	If rsEmp("emp_sawo_date") = "1900-01-01" Then
		emp_sawo_date = ""
	Else
		emp_sawo_date = rsEmp("emp_sawo_date")
	End If

	emp_email = rsEmp("emp_email")&"@k-one.co.kr"
%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_sex%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_person1")%>-<%=rsEmp("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsEmp("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsEmp("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsEmp("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("org_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("org_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("org_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("org_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("org_reside_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=stay_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_gunsok_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_end_gisan")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_yuncha_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_org_baldate")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_grade_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_birthday%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_jikmu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_last_edu")%></div></td>
    <td width="350">
		<div align="center" class="style1">
			<%=rsEmp("emp_family_zip")%>&nbsp;<%=rsEmp("emp_family_sido")%>&nbsp;<%=rsEmp("emp_family_gugun")%>&nbsp;<%=rsEmp("emp_family_dong")%>&nbsp;<%=rsEmp("emp_family_addr")%>
		</div>
	</td>
    <td width="350">
		<div align="center" class="style1">
			<%=rsEmp("emp_zipcode")%>&nbsp;<%=rsEmp("emp_sido")%>&nbsp;<%=rsEmp("emp_gugun")%>&nbsp;<%=rsEmp("emp_dong")%>&nbsp;<%=rsEmp("emp_addr")%>
		</div>
	</td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_tel_ddd")%>-<%=rsEmp("emp_tel_no1")%>-<%=rsEmp("emp_tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_hp_ddd")%>-<%=rsEmp("emp_hp_no1")%>-<%=rsEmp("emp_hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_email%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_emergency_tel")%></div></td>
    <%
	If rsEmp("emp_sawo_id") = "Y" Then
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
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_disabled")%>&nbsp;<%=rsEmp("emp_disab_grade")%></div></td>
    <td width="145">
		<div align="center" class="style1">
			<%=rsEmp("emp_military_id")%>&nbsp;<%=emp_military_date1%>&nbsp;<%=emp_military_date2%>&nbsp;<%=rsEmp("emp_military_grade")%>&nbsp;<%=rsEmp("emp_military_comm")%>
		</div>
	</td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_hobby")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_faith")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_marry_date%></div></td>
  </tr>
<%
	rsEmp.MoveNext()
Loop
Set rsStay = Nothing
rsEmp.Close() : Set rsEmp = Nothing
DBConn.Close() : Set DBConn = Nothing
%>
</table>
</body>
</html>