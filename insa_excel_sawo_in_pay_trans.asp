<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_emp
Dim sawo_empno
Dim sawo_emp_name

view_condi=Request("view_condi")
pmg_yymm=Request("pmg_yymm")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = pmg_yymm + "월 경조금 급여공제 현황 " + view_condi + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
         Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) ORDER BY de_company,de_emp_no ASC"
   else
         Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) and (de_company = '"+view_condi+"') ORDER BY de_company,de_emp_no ASC"
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
    <td colspan="8" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=pmg_yymm%>월&nbsp;경조금 급여공제 현황&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">년월</div></td>
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성  명</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">경조금</div></td>
  </tr>
    <%
		do until rs.eof 		
		de_emp_no = rs("de_emp_no")
		de_emp_name = rs("de_emp_name")
		
        if de_emp_no <> "" then
		   Sql="select * from emp_master where emp_no = '"&de_emp_no&"'"
		   Rs_emp.Open Sql, Dbconn, 1

		  if not Rs_emp.eof then
             emp_grade = Rs_emp("emp_grade")
			 emp_position = Rs_emp("emp_position")
		  end if
		  Rs_emp.Close()
		end if		
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("de_yymm")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("de_emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("de_emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_position%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("de_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("de_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(rs("de_sawo_amt"),0)%></div></td>
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
