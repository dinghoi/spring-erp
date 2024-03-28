<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_emp
Dim in_empno
Dim in_emp_name

curr_date = datevalue(mid(cstr(now()),1,10))

view_condi=Request("view_condi")
from_date=request("from_date")
to_date=request("to_date")

title_line = cstr(from_date) + " ~ " + cstr(to_date) + " " + " 경조금 신청 현황"

savefilename = title_line +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY ask_company,ask_date,ask_empno ASC"
where_sql = " WHERE (ask_company_process = '0') and (ask_company = '"+view_condi+"') and (ask_date > '"+from_date+"') and (ask_date < '"+to_date+"')"

sql = "select * from emp_sawo_ask " + where_sql + order_sql
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
    <td><div align="center" class="style1">경조일시</div></td>
    <td><div align="center" class="style1">경조구분</div></td>
    <td><div align="center" class="style1">경조유형</div></td>
    <td><div align="center" class="style1">경조장소</div></td>
    <td><div align="center" class="style1">기타내역</div></td>
  </tr>
    <%
		do until rs.eof 
		
		ask_empno = rs("ask_empno")
		ask_emp_name = rs("ask_emp_name")
		
        if ask_empno <> "" then
		   Sql="select * from emp_master where emp_no = '"&ask_empno&"'"
		   Rs_emp.Open Sql, Dbconn, 1

		  if not Rs_emp.eof then
             emp_grade = Rs_emp("emp_grade")
			 emp_position = Rs_emp("emp_position")
		  end if
		  Rs_emp.Close()
		end if		

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("ask_empno")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ask_emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_position%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ask_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ask_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("ask_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("ask_id")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ask_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ask_sawo_place")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("ask_sawo_comm")%></div></td>
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
