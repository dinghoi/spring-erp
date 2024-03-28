<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = view_condi + "인원현황" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_reside_place,emp_no,emp_in_date ASC"
where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"')"

sql = "select * from emp_master " + where_sql + order_sql
Rs.Open Sql, Dbconn, 1

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
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 
		
		stay_name = ""
		stay_name = rs("emp_stay_name")
		'stay_code = rs("emp_stay_code")
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
        
		emp_person2 = rs("emp_person2")
        if emp_person2 <> "" then
	        sex_id = mid(cstr(emp_person2),1,1)
	        if sex_id = "1" then
	                emp_sex = "남"
	        	  else
		            emp_sex = "여"
	        end if
	    end if
		if rs("emp_org_baldate") = "1900-01-01" then
		   emp_org_baldate = ""
		   else 
		   emp_org_baldate = rs("emp_org_baldate")
		end if
		if rs("emp_grade_date") = "1900-01-01" then
		   emp_grade_date = ""
		   else 
		   emp_grade_date = rs("emp_grade_date")
		end if
		
		emp_email = rs("emp_email") + "@k-won.co.kr"

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_sex%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_person1")%>-<%=rs("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_reside_place")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("cost_center")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
    <td width="145"><div align="center" class="style1"><%=stay_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_gunsok_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_end_gisan")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_yuncha_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_org_baldate")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_grade_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_birthday")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_jikmu")%></div></td>

    <td width="350"><div align="center" class="style1"><%=rs("emp_family_zip")%>&nbsp;<%=rs("emp_family_sido")%>&nbsp;<%=rs("emp_family_gugun")%>&nbsp;<%=rs("emp_family_dong")%>&nbsp;<%=rs("emp_family_addr")%></div></td>

    <td width="350"><div align="center" class="style1"><%=rs("emp_zipcode")%>&nbsp;<%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_tel_ddd")%>-<%=rs("emp_tel_no1")%>-<%=rs("emp_tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_email%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_emergency_tel")%></div></td>
    <%
	if rs("emp_sawo_id") = "Y" then
	   emp_sawo_id = "가입"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%>-<%=rs("emp_sawo_date")%></div></td>
    <% 
	   else
	   emp_sawo_id = "안함"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%></div></td>
    <%
	end if
	%>
    <td width="145"><div align="center" class="style1"><%=rs("emp_disabled")%>&nbsp;<%=rs("emp_disab_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_military_id")%>&nbsp;<%=rs("emp_military_date1")%>&nbsp;<%=rs("emp_military_date2")%>&nbsp;<%=rs("emp_military_grade")%>&nbsp;<%=rs("emp_military_comm")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_hobby")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_faith")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_marry_date")%></div></td>
  </tr>
	<%
	Rs.MoveNext()
	loop
	%>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
