<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

to_date = request("to_date")
in_grade = request("in_grade")  
in_company = request("in_company")  

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

if in_company = "" then
	in_company = "케이원정보통신"
	to_date = curr_date
	in_grade = "대리2"
end if

savefilename = "승진대상자 현황 -- "+ in_company +""+ in_grade +"" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
'Set rs_last = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if in_company = "" then
	in_company = "케이원정보통신"
	to_date = curr_date
	in_grade = "대리2급"
end if

if in_grade = "대리2급" then
	condi_sql = "emp_grade like '%사원%' and "
end if
if in_grade = "대리1급" then
	condi_sql = "emp_grade like '%대리2급%' and "
end if
if in_grade = "과장" then
	condi_sql = "(emp_grade like '%대리2급%') or (emp_grade like '%대리1급%') and "
end if
if in_grade = "차장" then
	'condi_sql = "emp_grade and '과장' and "
	condi_sql = "emp_grade like '%과장%' and "
end if
if in_grade = "부장" then
	condi_sql = "emp_grade like '%차장%' and "
end if

target_date = to_date

Sql = "SELECT * FROM emp_master where "+condi_sql+"isNull(emp_end_date) or emp_end_date = '1900-01-01' ORDER BY emp_first_date,emp_no DESC"
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;승진대상자 현황>&nbsp;(<%=in_company%>)&nbsp;<%=in_grade%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">생년월일</div></td>
    <td><div align="center" class="style1">현직급</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">최초입사일</div></td>
    <td><div align="center" class="style1">입사일</div></td>
    <td><div align="center" class="style1">최종승진일</div></td>
    <td><div align="center" class="style1">대상년한</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">본부</div></td>
    <td><div align="center" class="style1">사업부</div></td>
    <td><div align="center" class="style1">팀</div></td>
  </tr>
    <%
		do until rs.eof 
		
		if rs("emp_grade_date") = "1900-01-01" then
		   emp_grade_date = ""
		   else 
		   emp_grade_date = rs("emp_grade_date")
		end if

        if emp_grade_date <> "" then 
			   year_cnt = datediff("yyyy", rs("emp_grade_date"), target_date)
               mon_cnt = datediff("m", rs("emp_grade_date"), target_date)
               day_cnt = datediff("d", rs("emp_grade_date"), target_date) 
			else 
			   year_cnt = datediff("yyyy", rs("emp_first_date"), target_date)
               mon_cnt = datediff("m", rs("emp_first_date"), target_date)
               day_cnt = datediff("d", rs("emp_first_date"), target_date) 
		end if
				
		target_cnt = cint(mon_cnt)		
		if (in_grade = "대리2급" or in_grade = "대리1급") and target_cnt > 24 then
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_birthday")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_grade_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=mon_cnt%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
  </tr>
	<%
		    else if in_grade = "과장" and Rs("emp_grade") = "대리1급" and target_cnt > 36 then
	%>	    
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_birthday")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_grade_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=mon_cnt%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
  </tr>
	<%    
			    else if in_grade = "과장" and Rs("emp_grade") = "대리2급" and target_cnt > 36 then
	%>	    
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_birthday")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_grade_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=mon_cnt%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
  </tr>
	<%    
	              end if
	        end if
	end if
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
