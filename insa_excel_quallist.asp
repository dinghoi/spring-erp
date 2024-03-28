<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_company=Request("view_company")
view_condi=Request("view_condi")
condi = Request("condi")

if view_condi = "전체" then
	condi = ""
end if

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "자격증 보유현황 -- "+ condi +""+ view_condi +"" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "상주처회사" then
'            Sql= "select * " & _
'	               "    from emp_qual a, emp_master b " & _
'	               "    where a.qual_empno = b.emp_no AND b.emp_reside_company like '%" + condi + "%' " & _
'				   "    ORDER BY qual_empno ASC"
				   
			Sql= "select * " & _
	               "    from emp_qual a, emp_master b " & _
	               "    where a.qual_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_company = '"&view_company&"') and (b.emp_reside_company like '%" + condi + "%') " & _
				   "    ORDER BY qual_empno ASC"  
		    Rs.Open Sql, Dbconn, 1
end if
if view_condi = "자격증명" then
'	condi_sql = " where qual_type like '%" + condi + "%'"
'	Sql = "SELECT * FROM emp_qual "+condi_sql+" ORDER BY qual_empno ASC"
	
	Sql= "select * " & _
	               "    from emp_qual a, emp_master b " & _
	               "    where a.qual_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_company = '"&view_company&"') and (a.qual_type like '%" + condi + "%') " & _
				   "    ORDER BY qual_empno ASC" 
    Rs.Open Sql, Dbconn, 1
end if
if view_condi = "전체" then
'	condi_sql = ""
'	Sql = "SELECT * FROM emp_qual "+condi_sql+" ORDER BY qual_empno ASC"
	
	Sql= "select * " & _
	               "    from emp_qual a, emp_master b " & _
	               "    where a.qual_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_company = '"&view_company&"') " & _
				   "    ORDER BY qual_empno ASC"
    Rs.Open Sql, Dbconn, 1
end if

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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;자격증 보유현황>&nbsp;(<%=condi%>)&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">자격종목</div></td>
    <td><div align="center" class="style1">등급</div></td>
    <td><div align="center" class="style1">발급기관</div></td>
    <td><div align="center" class="style1">자격등록번호</div></td>
    <td><div align="center" class="style1">취득일</div></td>
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">주민등록번호</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">상주처회사</div></td>
  </tr>
    <%
		do until rs.eof 
		
        qual_empno = rs("qual_empno")
        if qual_empno <> "" then
	       Sql="select * from emp_master where emp_no = '"&qual_empno&"'"
	       Rs_emp.Open Sql, Dbconn, 1

	       if not Rs_emp.eof then
              emp_name = Rs_emp("emp_name")
	    	  emp_grade = Rs_emp("emp_grade")
			  emp_job = Rs_emp("emp_job")
	          emp_position = Rs_emp("emp_position")
			  emp_org_code = Rs_emp("emp_org_code")
			  emp_org_name = Rs_emp("emp_org_name")
	          emp_company = Rs_emp("emp_company")
			  emp_team = Rs_emp("emp_team")
			  emp_reside_place = Rs_emp("emp_reside_place")
			  emp_reside_company = Rs_emp("emp_reside_company")
			  emp_person1 = Rs_emp("emp_person1")
			  emp_person2 = Rs_emp("emp_person2")
		   end if
	       Rs_emp.Close()
	    end if	

	%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="left" class="style1"><%=rs("qual_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("qual_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("qual_org")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rs("qual_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("qual_pass_date")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("qual_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_person1%>-<%=emp_person2%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_job%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_company%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_team%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_org_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_reside_company%></div></td>
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
