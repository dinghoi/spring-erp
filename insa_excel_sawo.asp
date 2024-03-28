<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_emp
Dim sawo_empno
Dim sawo_emp_name

view_condi=Request("view_condi")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "경조회 가입 현황 " + view_condi + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY sawo_company,sawo_org_name,sawo_date,sawo_empno ASC"
if view_condi = "전체" then
         where_sql = " WHERE sawo_target = 'Y' or sawo_target = 'N'"
   else
         where_sql = " WHERE sawo_company = '"+view_condi+"' and sawo_target = 'Y' or sawo_target = 'N'"
end if

sql = "select * from emp_sawo_mem " + where_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=view_condi%> &nbsp;경조회 가입현황</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성  명</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">가입일</div></td>
    <td><div align="center" class="style1">가입구분</div></td>
    <td><div align="center" class="style1">탈퇴일</div></td>
    <td><div align="center" class="style1">탈퇴구분</div></td>
    <td><div align="center" class="style1">급여공제</div></td>
    <td><div align="center" class="style1">납입횟수</div></td>
    <td><div align="center" class="style1">납입금액</div></td>
    <td><div align="center" class="style1">지급횟수</div></td>
    <td><div align="center" class="style1">지급금액</div></td>
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
    <td width="115"><div align="center" class="style1"><%=rs("sawo_empno")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_position%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("sawo_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("sawo_date")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("sawo_id")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("sawo_out_date")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("sawo_out")%></div></td>
    <% If rs("sawo_target") = "Y" then sawo_target = "공제" end if %>
    <% If rs("sawo_target") = "N" then sawo_target = "안함" end if %>
    <td width="59"><div align="center" class="style1"><%=sawo_target%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(rs("sawo_in_count"),0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(rs("sawo_in_pay"),0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(rs("sawo_give_count"),0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=formatnumber(rs("sawo_give_pay"),0)%></div></td>
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
