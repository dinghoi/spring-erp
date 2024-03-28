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
Dim view_condi, condi, curr_date, condi_sql, savefilename
Dim rsReport

view_condi = Request("view_condi")
condi = Request("condi")

curr_date = datevalue(mid(cstr(now()),1,10))

if view_condi = "" then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
end if

savefilename = "직원현황 -- "& condi &""& view_condi &"" & cstr(curr_date) & ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

if view_condi = "" then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
end if

if view_condi = "소속조직별" then
	'condi_sql = "(emp_org_name like '%" + condi + "%') and "
	condi_sql = "AND eomt.org_name LIKE '%" & condi & "%' "
end if
if view_condi = "성명" then
	condi_sql = "AND emtt.emp_name LIKE '%" & condi & "%' "
end if
if view_condi = "직급별" then
	condi_sql = "AND emtt.emp_grade LIKE '%" & condi & "%' "
end if
if view_condi = "직위별" then
	condi_sql = "AND emtt.emp_job LIKE '%" & condi & "%' "
end if
if view_condi = "직책별" then
	condi_sql = "AND emtt.emp_position LIKE '%" & condi & "%' "
end if
if view_condi = "회사별" then
	'condi_sql = "(emp_company like '%" + condi + "%') and "
	condi_sql = "AND eomt.org_company LIKE '%" & condi & "%' "
end if
if view_condi = "본부별" then
	'condi_sql = "(emp_bonbu like '%" + condi + "%') and "
	condi_sql = "AND eomt.org_bonbu LIKE '%" & condi & "%' "
end if
if view_condi = "사업부별" then
	'condi_sql = "(emp_saupbu like '%" + condi + "%') and "
	condi_sql = "AND eomt.org_saupbu LIKE '%" & condi & "%' "
end if
if view_condi = "팀별" then
	'condi_sql = "(emp_team like '%" + condi + "%') and "
	condi_sql = "AND eomt.org_team LIKE '%" & condi & "%' "
end if
if view_condi = "상주처 회사별" then
	'condi_sql = "(emp_reside_company like '%" + condi + "%') and "
	condi_sql = "AND eomt.org_reside_company LIKE '%" & condi & "%' "
end if
if view_condi = "입사일자별" then
	condi_sql = "AND emp_in_date LIKE '%" & condi & "%' "
end if

'Sql = "SELECT * FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY objBuilder.Append "SELECT emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_birthday, emtt.emp_no, "
objBuilder.Append "SELECT emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_birthday, emtt.emp_no, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_in_date, "
objBuilder.Append "	emtt.emp_org_name, emtt.emp_first_date, emtt.emp_reside_place, emtt.emp_company, "
objBuilder.Append "	emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, emtt.emp_stay_code, emtt.emp_person2, "
objBuilder.Append "	emtt.emp_military_date2, emtt.emp_marry_date, emtt.emp_grade_date, emtt.emp_email, "
objBuilder.Append "	emtt.emp_military_date1, emtt.emp_end_date, emtt.emp_org_baldate, emtt.emp_sawo_date, "
objBuilder.Append "	emtt.cost_center, emtt.emp_type, emtt.emp_person1, emtt.emp_person2, emtt.emp_gunsok_date, "
objBuilder.Append "	emtt.emp_end_gisan, emtt.emp_yuncha_date, emtt.emp_jikmu, emtt.emp_last_edu, "
objBuilder.Append "	emtt.emp_family_sido, emtt.emp_family_gugun, emtt.emp_family_addr, emtt.emp_family_dong, "
objBuilder.Append "	emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, "
objBuilder.Append "	emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_hp_ddd, emtt.emp_hp_no1, "
objBuilder.Append "	emtt.emp_hp_no2, emtt.emp_emergency_tel, emtt.emp_sawo_id, emtt.emp_disabled, emtt.emp_disab_grade, "
objBuilder.Append "	emtt.emp_military_id, emtt.emp_military_grade, emtt.emp_military_comm, "
objBuilder.Append "	emtt.emp_hobby, emtt.emp_faith, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_reside_place "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append condi_sql
objBuilder.Append "ORDER BY emtt.emp_no, emtt.emp_name ASC "

Set rsReport = Server.CreateObject("ADODB.RecordSet")
rsReport.Open objBuilder.ToString(), Dbconn, 1
objBuilder.Clear()

%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE HTML>
<html lang="ko">
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;직원 현황>&nbsp;(<%=condi%>)&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">성별</div></td>
    <td><div align="center" class="style1">직원구분</div></td>
    <td><div align="center" class="style1">코스트센타</div></td>
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
	Dim stay_name, stay_code, rs_stay, emp_person2, sex_id, emp_sex
	Dim emp_birthday, emp_military_date1, emp_military_date2, emp_marry_date
	Dim emp_grade_date, emp_end_date, emp_org_baldate, emp_sawo_date, emp_email
	Dim emp_sawo_id

	do until rsReport.EOF

		stay_name = ""
		stay_code = rsReport("emp_stay_code")

        if stay_code <> "" then
		   'Sql="select * from emp_stay where stay_code = '"&stay_code&"'"
		   objBuilder.Append "SELECT stay_name FROM emp_stay WHERE stay_code = '"&stay_code&"' "

		   Set rs_stay = DBConn.Execute(objBuilder.ToString())

		  'do until rs_stay.eof
		  if not rs_stay.eof then
             stay_name = rs_stay("stay_name")
	         'rs_stay.movenext()
			 'loop
		  end if
		  rs_stay.Close()
		end if

		emp_person2 = rsReport("emp_person2")
		if emp_person2 <> "" then
		   sex_id = mid(cstr(emp_person2),1,1)
		   if sex_id = "1" then
				 emp_sex = "남"
		   else
				 emp_sex = "여"
		   end if
		end if

		if rsReport("emp_birthday") = "1900-01-01" then
			   emp_birthday = ""
		else
			   emp_birthday = rsReport("emp_birthday")
		end if
		if rsReport("emp_military_date1") = "1900-01-01" then
			   emp_military_date1 = ""
			   emp_military_date2 = ""
		else
			   emp_military_date1 = rsReport("emp_military_date1")
			   emp_military_date2 = rsReport("emp_military_date2")
		end if
		if rsReport("emp_marry_date") = "1900-01-01" then
			   emp_marry_date = ""
			else
			   emp_marry_date = rsReport("emp_marry_date")
		end if
		if rsReport("emp_grade_date") = "1900-01-01" then
			   emp_grade_date = ""
		   else
			   emp_grade_date = rsReport("emp_grade_date")
		end if
		if rsReport("emp_end_date") = "1900-01-01" then
			   emp_end_date = ""
			else
			   emp_end_date = rsReport("emp_end_date")
		end if
		if rsReport("emp_org_baldate") = "1900-01-01" then
			   emp_org_baldate = ""
		   else
			   emp_org_baldate = rsReport("emp_org_baldate")
		end if
		if rsReport("emp_sawo_date") = "1900-01-01" then
			   emp_sawo_date = ""
		   else
			   emp_sawo_date = rsReport("emp_sawo_date")
		end if

		emp_email = rsReport("emp_email") + "@k-won.co.kr"

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_sex%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("cost_center")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_person1")%>-<%=rsReport("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsReport("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsReport("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rsReport("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("org_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("org_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("org_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("org_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("org_reside_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=stay_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_gunsok_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_end_gisan")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_yuncha_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_org_baldate")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_grade_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_birthday%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsReport("emp_jikmu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_last_edu")%></div></td>

    <td width="350"><div align="left" class="style1"><%=rsReport("emp_family_sido")%>&nbsp;<%=rsReport("emp_family_gugun")%>&nbsp;<%=rsReport("emp_family_dong")%>&nbsp;<%=rsReport("emp_family_addr")%></div></td>

    <td width="350"><div align="left" class="style1"><%=rsReport("emp_sido")%>&nbsp;<%=rsReport("emp_gugun")%>&nbsp;<%=rsReport("emp_dong")%>&nbsp;<%=rsReport("emp_addr")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_tel_ddd")%>-<%=rsReport("emp_tel_no1")%>-<%=rsReport("emp_tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_hp_ddd")%>-<%=rsReport("emp_hp_no1")%>-<%=rsReport("emp_hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_email%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_emergency_tel")%></div></td>
    <% 'response.write(rsReport("emp_stay_code"))
	   'response.End %>
    <%
	if rsReport("emp_sawo_id") = "Y" then
	   emp_sawo_id = "가입"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%>-<%=emp_sawo_date%></div></td>
    <%
	   else
	   emp_sawo_id = "안함"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%></div></td>
    <%
	end if
	%>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_disabled")%>&nbsp;<%=rsReport("emp_disab_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_military_id")%>&nbsp;<%=emp_military_date1%>&nbsp;<%=emp_military_date2%>&nbsp;<%=rsReport("emp_military_grade")%>&nbsp;<%=rsReport("emp_military_comm")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_hobby")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rsReport("emp_faith")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_marry_date%></div></td>
  </tr>
	<%
	rsReport.MoveNext()
	loop
	%>
</table>
</body>
</html>
<%
rsReport.Close() : Set rsReport = Nothing
DBConn.Close() : Set DBConn = Nothing
%>
