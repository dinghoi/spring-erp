<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_emp
Dim give_empno
Dim give_emp_name

view_condi=Request("view_condi")
ask_process=request("ask_process")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

if ask_process = "1" then 
   title_line = cstr(from_date) + " ~ " + cstr(to_date) + " " + " 경조금 지급현황"
   else
   title_line = cstr(from_date) + " ~ " + cstr(to_date) + " " + " 경조회 경조금 지급현황"
end if

savefilename = title_line +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY give_company,give_date,give_empno DESC"
if ask_process = "1" then 
   where_sql = " WHERE (give_ask_process = '"+ask_process+"') and (give_company = '"+view_condi+"') and (give_date > '"+from_date+"') and (give_date < '"+to_date+"')"
   else
   where_sql = " WHERE give_ask_process = '"+ask_process+"'"
end if

sql = "select * from emp_sawo_give " + where_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성  명</div></td>
    <td><div align="center" class="style1">현직급</div></td>
    <td><div align="center" class="style1">현직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">지급일</div></td>
    <td><div align="center" class="style1">지급구분</div></td>
    <td><div align="center" class="style1">지급유형</div></td>
    <td><div align="center" class="style1">발생일</div></td>
    <td><div align="center" class="style1">지급금액</div></td>
    <td><div align="center" class="style1">경조장소</div></td>
    <td><div align="center" class="style1">경조내용</div></td>
    <td><div align="center" class="style1">비  고</div></td>
  </tr>
    <%
		do until rs.eof 
		
		give_empno = rs("give_empno")
		give_emp_name = rs("give_emp_name")
		
        if give_empno <> "" then
		   Sql="select * from emp_master where emp_no = '"&give_empno&"'"
		   Rs_emp.Open Sql, Dbconn, 1

		  if not Rs_emp.eof then
             emp_grade = Rs_emp("emp_grade")
			 emp_position = Rs_emp("emp_position")
		  end if
		  Rs_emp.Close()
		end if		

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("give_empno")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_position%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("give_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("give_id")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_sawo_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=formatnumber(clng(rs("give_pay")),0)%></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_sawo_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_sawo_comm")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("give_comment")%></div></td>
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
