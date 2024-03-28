<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")

de_sawo_amt = 7000

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "경조회 경조회비(공제) -- "+ view_condi +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   Sql = "select * from emp_sawo_mem where sawo_out = '' or isnull(sawo_out) ORDER BY sawo_empno ASC"
   else  
   Sql = "select * from emp_sawo_mem where sawo_company = '"+view_condi+"' and (sawo_out = '' or isnull(sawo_out)) ORDER BY sawo_empno ASC"
end if
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
    <td colspan="11" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;경조회 경조회비(공제)--<%=view_condi%>&nbsp;</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">귀속년월</div></td>
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">가입일</div></td>
    <td><div align="center" class="style1">공제금액</div></td>
  </tr>
    <%
		do until rs.eof 
		
		   sawo_empno = rs("sawo_empno")
		   sawo_emp_name = rs("sawo_emp_name")
		
           if sawo_empno <> "" then
		          Sql="select * from emp_master where emp_no = '"&sawo_empno&"'"
		          Rs_emp.Open Sql, Dbconn, 1

		          if not Rs_emp.eof then
                         emp_grade = Rs_emp("emp_grade")
		                 emp_position = Rs_emp("emp_position")
		          end if
	              Rs_emp.Close()
	        end if		

	%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_empno")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_emp_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_position%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=formatnumber(de_sawo_amt,0)%></div></td>
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
