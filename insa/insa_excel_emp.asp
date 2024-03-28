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
Dim rsEmp

view_condi = Request("view_condi")
curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = view_condi & "인원현황" & cstr(curr_date) & ".xls"

Call ViewExcelType(savefilename)

objBuilder.Append "SELECT emtt.emp_stay_name, emtt.emp_person2, emtt.emp_org_baldate, emtt.emp_grade_date, "
objBuilder.Append "	emtt.emp_email, emtt.emp_no, emtt.emp_name, emtt.emp_type, emtt.emp_person1, emtt.emp_person2, "
objBuilder.Append "	emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_org_name, emtt.emp_company, emtt.emp_bonbu, "
objBuilder.Append "	emtt.emp_saupbu, emtt.emp_team, emtt.emp_reside_place, emtt.cost_center, emtt.emp_first_date, "
objBuilder.Append "	emtt.emp_in_date, emtt.emp_gunsok_date, emtt.emp_end_gisan, emtt.emp_yuncha_date, emtt.emp_org_baldate, "
objBuilder.Append "	emtt.emp_grade_date, emtt.emp_birthday, emtt.emp_jikmu, emtt.emp_family_zip, emtt.emp_family_sido, "
objBuilder.Append "	emtt.emp_family_gugun, emtt.emp_family_dong, emtt.emp_family_addr, emtt.emp_sido, emtt.emp_gugun, "
objBuilder.Append "	emtt.emp_dong, emtt.emp_addr, emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_hp_ddd, "
objBuilder.Append "	emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_emergency_tel, emtt.emp_sawo_id, emtt.emp_sawo_date, "
objBuilder.Append "	emtt.emp_disabled, emtt.emp_disab_grade, emtt.emp_military_id, emtt.emp_military_date1, "
objBuilder.Append "	emtt.emp_military_date2, emtt.emp_military_grade, emtt.emp_military_comm, emtt.emp_hobby, "
objBuilder.Append "	emtt.emp_faith, emtt.emp_marry_date, emtt.emp_zipcode, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_reside_place, "
objBuilder.Append "	dpit.dz_id "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "INNER JOIN memb AS mem ON emtt.emp_no = mem.user_id AND mem.grade < '6' "
objBuilder.Append "LEFT OUTER JOIN dz_pay_info AS dpit ON emtt.emp_no = dpit.emp_no "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emtt.emp_no < '900000' "

If view_condi <> "전체" Then
	objBuilder.Append "	AND eomt.org_company = '"&view_condi&"' "
End If

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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=view_condi%> &nbsp;인원 현황&nbsp;<%=curr_date%></div></td>
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
    <td><div align="center" class="style1">비용구분</div></td>
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
	<td><div align="center" class="style1">급여ID</div></td>
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
<%
Dim stay_name, emp_person2, emp_org_baldate, emp_email, sex_id
Dim emp_sex, emp_grade_date, emp_sawo_id

do until rsEmp.EOF

	stay_name = ""
	stay_name = rsEmp("emp_stay_name")
	'stay_code = rsEmp("emp_stay_code")
	'if stay_code <> "" then
	'   Sql="select * from emp_stay where stay_code = '"&stay_code&"'"
	'   Rs_stay.Open Sql, Dbconn, 1

	'  'do until rs_stay.eof
	'  if not rs_stay.eof then
	'     stay_name = rs_stay("stay_name")
	'     'rs_stay.movenext()
	'	 'loop
	'  end if
	'  rs_stay.Close()
	'end if

	emp_person2 = rsEmp("emp_person2")
	if emp_person2 <> "" then
		sex_id = mid(cstr(emp_person2),1,1)
		if sex_id = "1" then
				emp_sex = "남"
			  else
				emp_sex = "여"
		end if
	end If

	if rsEmp("emp_org_baldate") = "1900-01-01" then
	   emp_org_baldate = ""
	   else
	   emp_org_baldate = rsEmp("emp_org_baldate")
	end If

	if rsEmp("emp_grade_date") = "1900-01-01" then
	   emp_grade_date = ""
	   else
	   emp_grade_date = rsEmp("emp_grade_date")
	end if

	emp_email = rsEmp("emp_email") & "@k-won.co.kr"
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
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_reside_place")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("cost_center")%></div></td>
    <% 'response.write(rsEmp("emp_stay_code"))
	   'response.End %>
    <td width="145"><div align="center" class="style1"><%=stay_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_gunsok_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_end_gisan")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_yuncha_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_org_baldate")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_grade_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_birthday")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsEmp("emp_jikmu")%></div></td>

    <td width="350"><div align="center" class="style1"><%=rsEmp("emp_family_zip")%>&nbsp;<%=rsEmp("emp_family_sido")%>&nbsp;<%=rsEmp("emp_family_gugun")%>&nbsp;<%=rsEmp("emp_family_dong")%>&nbsp;<%=rsEmp("emp_family_addr")%></div></td>

    <td width="350"><div align="center" class="style1"><%=rsEmp("emp_zipcode")%>&nbsp;<%=rsEmp("emp_sido")%>&nbsp;<%=rsEmp("emp_gugun")%>&nbsp;<%=rsEmp("emp_dong")%>&nbsp;<%=rsEmp("emp_addr")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_tel_ddd")%>-<%=rsEmp("emp_tel_no1")%>-<%=rsEmp("emp_tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_hp_ddd")%>-<%=rsEmp("emp_hp_no1")%>-<%=rsEmp("emp_hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_email%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_emergency_tel")%></div></td>
    <%
	if rsEmp("emp_sawo_id") = "Y" then
	   emp_sawo_id = "가입"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%>-<%=rsEmp("emp_sawo_date")%></div></td>
    <%
	else
	   emp_sawo_id = "안함"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%></div></td>
    <%
	end if
	%>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_disabled")%>&nbsp;<%=rsEmp("emp_disab_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_military_id")%>&nbsp;<%=rsEmp("emp_military_date1")%>&nbsp;<%=rsEmp("emp_military_date2")%>&nbsp;<%=rsEmp("emp_military_grade")%>&nbsp;<%=rsEmp("emp_military_comm")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_hobby")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_faith")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsEmp("emp_marry_date")%></div></td>
	<td width="145"><div align="center" class="style1"><%=rsEmp("dz_id")%></div></td>
  </tr>
	<%
		rsEmp.MoveNext()
	Loop
	rsEmp.Close() : Set rsEmp = Nothing
	DBConn.Close() : Set DBConn = Nothing
	%>
</table>
</body>
</html>