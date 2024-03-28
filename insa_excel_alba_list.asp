<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = view_condi + "사업소득자 현황" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY company,org_name,draft_no ASC"
where_sql = " WHERE end_yn <> 'Y' and (company = '"&view_condi&"')"

sql = "select * from emp_alba_mst " + where_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=view_condi%> &nbsp;사업소득자 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">등록번호</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">업무등록일</div></td>
    <td><div align="center" class="style1">소득구분</div></td>
    <td><div align="center" class="style1">주민번호</div></td>
    <td><div align="center" class="style1">내/외국인</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">본부</div></td>
    <td><div align="center" class="style1">사업부</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">비용회사</div></td>
    <td><div align="center" class="style1">전자결재번호</div></td>
    <td><div align="center" class="style1">기안일자</div></td>
    <td><div align="center" class="style1">기안자</div></td>
    <td><div align="center" class="style1">전화번호</div></td> 
    <td><div align="center" class="style1">핸드폰</div></td>
    <td><div align="center" class="style1">e메일</div></td>
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("draft_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("draft_man")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("draft_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("draft_tax_id")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("person_no1")%>-<%=rs("person_no2")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("nation_id")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("org_name")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
    <td width="115"><div align="center" class="style1"><%=rs("cost_company")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("sign_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("deposit_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("deposit_man")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("tel_ddd")%>-<%=rs("tel_no1")%>-<%=rs("tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("hp_ddd")%>-<%=rs("hp_no1")%>-<%=rs("hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("e_mail")%></div></td>
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
