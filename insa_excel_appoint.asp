<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
app_id=Request("app_id")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = view_condi + "(" + app_id + ") - 인사발령 현황(" + from_date + " ∼ " + to_date + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   if app_id = "전체" then
           Sql = "select * from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	  else 
		   Sql = "select * from emp_appoint where app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
   end if	   
   else  
      if app_id = "전체" then
	          Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
		 else	  
			  Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	  end if
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
    <td colspan="12" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=from_date%>&nbsp;∼&nbsp;<%=to_date%> &nbsp;인사발령 현황>&nbsp;(<%=view_condi%>)</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
    <td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성명</div></td>
    <td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령일</div></td>
    <td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령구분</div></td>
    <td rowspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">발령유형</div></td>
    <td colspan="3" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">발령전</div></td>
    <td colspan="4" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">발령후</div></td>
  </tr>
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">소속</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">직급/책</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">소속</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">직급/책</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">발령내용</div></td>
  </tr>  
    <%
	  do until rs.eof 
		
	%>
  <tr valign="middle" class="style11">
    <td width="95"><div align="center" class="style1"><%=rs("app_empno")%></div></td>
    <td width="95"><div align="center" class="style1"><%=rs("app_emp_name")%></div></td>
    <td width="95"><div align="center" class="style1"><%=rs("app_date")%></div></td>
    <td width="95"><div align="center" class="style1"><%=rs("app_id")%></div></td>
    <td width="95"><div align="center" class="style1"><%=rs("app_id_type")%></div></td>
    <td width="95"><div align="center" class="style1"><%=rs("app_to_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("app_to_org")%>(<%=rs("app_to_orgcode")%>)</div></td>
    <td width="145"><div align="center" class="style1"><%=rs("app_to_grade")%>-<%=rs("app_to_position")%></div></td>
    <td width="95"><div align="center" class="style1"><%=rs("app_be_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("app_be_org")%>(<%=rs("app_be_orgcode")%>)</div></td>
    <td width="145"><div align="center" class="style1"><%=rs("app_be_grade")%>-<%=rs("app_be_position")%></div></td>
    <td width="300" class="left"><div align="center" class="style1"><%=rs("app_start_date")%>&nbsp;-&nbsp;<%=rs("app_finish_date")%>&nbsp;<%=rs("app_be_enddate")%>&nbsp;<%=rs("app_reward")%>&nbsp;:&nbsp;<%=rs("app_comment")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
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
