<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

owner_view=Request("owner_view")
field_check=Request("field_check")
field_view=Request("field_view")
	
curr_date = datevalue(mid(cstr(now()),1,10))

if owner_view = "C" then
	owner_gubun = "회사 "
  elseif owner_view = "P" then
	owner_gubun = "개인 "
  else  
  	owner_gubun = "전체"
end if

savefilename = owner_gubun + " 차량 현황" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * FROM car_info "

if owner_view = "C" then
	owner_sql = " where car_owner = '회사' "
  elseif owner_view = "P" then
	owner_sql = " where car_owner = '개인' "
  else  
  	owner_sql = " where (car_owner = '개인' or car_owner = '회사') "
end if

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

order_sql = " ORDER BY car_no DESC"

sql = base_sql + owner_sql + field_sql + order_sql
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=owner_gubun%> &nbsp;차량 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">차량번호</div></td>
    <td><div align="center" class="style1">차종</div></td>
    <td><div align="center" class="style1">연식</div></td>
    <td><div align="center" class="style1">유류종류</div></td>
    <td><div align="center" class="style1">차량소유회사</div></td>
    <td><div align="center" class="style1">사용부서</div></td>
    <td><div align="center" class="style1">용도</div></td>
    <td><div align="center" class="style1">운행자</div></td>
    <td><div align="center" class="style1">차량등록일</div></td>
    <td><div align="center" class="style1">운행Km</div></td>
    <td><div align="center" class="style1">보험기간</div></td>
    <td><div align="center" class="style1">보험회사</div></td>
    <td><div align="center" class="style1">보험료</div></td>
    <td><div align="center" class="style1">최종점검일</div></td>
    <td><div align="center" class="style1">차량상태</div></td>
    <td><div align="center" class="style1">차량정보</div></td>
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("car_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("car_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_year")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("oil_kind")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_company")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_use_dept")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_use")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("owner_emp_name")%>(<%=rs("owner_emp_no")%>)&nbsp;</div></td>
    <td width="145"><div align="center" class="style1"><%=rs("car_reg_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=formatnumber(rs("last_km"),0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("insurance_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("insurance_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=formatnumber(rs("insurance_amt"),0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("last_check_date")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
    <td width="115"><div align="center" class="style1"><%=rs("car_status")%></div></td>
    <td width="200"><div align="center" class="style1"><%=rs("car_comment")%></div></td>
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
