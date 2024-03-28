<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

from_date=Request("from_date")
to_date=Request("to_date")
company = Request("company")
cfm_type = Request("cfm_type")

if company = "전체" then
	com_sql = ""
  else
  	com_sql = " (cfm_company ='"+company+"') and "
end if
if cfm_type = "전체" then
	type_sql = ""
  else
  	type_sql = " (cfm_type ='"+cfm_type+"') and "
end if

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "제증명 발급현황 -- "+ cfm_company +""+ cfm_type +"" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT * FROM emp_confirm where "+com_sql+type_sql+" cfm_date >= '"+from_date+"' and cfm_date <= '"+to_date+"' ORDER BY cfm_seq DESC"
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;제증명 발급현황>&nbsp;(<%=cfm_company%>)&nbsp;<%=cfm_type%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">주민등록번호</div></td>
    <td><div align="center" class="style1">직위</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">제증명</div></td>
    <td><div align="center" class="style1">발급일자</div></td>
    <td><div align="center" class="style1">용도</div></td>
    <td><div align="center" class="style1">사용처</div></td>
    <td><div align="center" class="style1">비고</div></td>
  </tr>
    <%
		do until rs.eof

        cfm_empno = rs("cfm_empno")
        if cfm_empno <> "" then
	       Sql="select * from emp_master where emp_no = '"&cfm_empno&"'"
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
    <td width="59"><div align="center" class="style1"><%=rs("cfm_empno")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("cfm_emp_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_person1")%>-<%=rs("cfm_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("cfm_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("cfm_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_use")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("cfm_use_dept")%></div></td>
    <td width="200"><div align="center" class="style1"><%=rs("cfm_comment")%></div></td>
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
