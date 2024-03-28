<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_emp
Dim in_empno
Dim in_emp_name

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "경조회 회비 납부 현황" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY in_company,in_date,in_empno ASC"
'where_sql = " WHERE sawo_target = 'Y' or sawo_target = 'N'"
where_sql = ""

sql = "select * from emp_sawo_in " + where_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;경조회 회비 납부현황</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성  명</div></td>
    <td><div align="center" class="style1">현직급</div></td>
    <td><div align="center" class="style1">현직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">납부일</div></td>
    <td><div align="center" class="style1">납부금액</div></td>
    <td><div align="center" class="style1">비  고</div></td>
  </tr>
    <%
		do until rs.eof 
		
		in_empno = rs("in_empno")
		in_emp_name = rs("in_emp_name")
		
        if in_empno <> "" then
		   Sql="select * from emp_master where emp_no = '"&in_empno&"'"
		   Rs_emp.Open Sql, Dbconn, 1

		  if not Rs_emp.eof then
             emp_grade = Rs_emp("emp_grade")
			 emp_position = Rs_emp("emp_position")
		  end if
		  Rs_emp.Close()
		end if		

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("in_empno")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("in_emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_position%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("in_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("in_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(clng(rs("in_pay")),0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("in_comment")%></div></td>
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
