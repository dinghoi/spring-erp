<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
condi = Request("condi")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "휴직발령 현황 -- "+ condi +""+ view_condi +"" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

view_sort = "ASC"

order_Sql = " ORDER BY app_date,app_empno,app_seq " + view_sort
where_sql = " WHERE app_id = '휴직발령' and app_bokjik_id = 'N'"
'where_sql = ""


sql = "select * from emp_appoint " + where_sql + order_sql 
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;휴직 발령 현황</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성  명</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">발령일</div></td>
    <td><div align="center" class="style1">휴직유형</div></td>
    <td><div align="center" class="style1">휴직기간</div></td>
    <td><div align="center" class="style1">휴직사유</div></td>
  </tr>
    <%
		do until rs.eof 

	    app_empno = rs("app_empno")
	    app_emp_name = rs("app_emp_name")
	    if app_empno <> "" then
	       Sql="select * from emp_master where emp_no = '"&app_empno&"'"
	       Rs_emp.Open Sql, Dbconn, 1

	       if not Rs_emp.eof then
           emp_grade = Rs_emp("emp_grade")
	       emp_grade = Rs_emp("emp_job")
		   emp_position = Rs_emp("emp_position")
		   emp_org_code = Rs_emp("emp_org_code")
		   emp_org_name = Rs_emp("emp_org_name")
		   emp_company = Rs_emp("emp_company")
		   end if
	       Rs_emp.Close()
	     end if		

	%>
  <tr valign="middle" class="style11">
    <td width="59"><div align="left" class="style1"><%=rs("app_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("app_emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="59"><div align="left" class="style1"><%=emp_position%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("app_to_company")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("app_to_org")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("app_date")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("app_id_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("app_start_date")%>-<%=rs("app_finish_date")%></div></td>
    <td width="200"><div align="center" class="style1"><%=rs("app_comment")%></div></td>
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
