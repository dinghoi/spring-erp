<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
from_date=request("from_date")
to_date=request("to_date")
pmg_yymm=request("pmg_yymm")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "야·특근 현황(수당) -- "+ view_condi +"(" + from_date + "∼" + to_date + ").xls"

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
   Sql = "select * from overtime where work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"
   else  
   Sql = "select * from overtime where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"
end if
Rs.Open Sql, Dbconn, 1
do until rs.eof
    overtime_count = overtime_count + 1
    sum_overtime_pay = sum_overtime_pay + int(rs("overtime_amt"))
	rs.movenext()
loop
rs.close()

if view_condi = "전체" then
   Sql = "select * from overtime where work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,bonbu,saupbu,team,org_name,work_date,mg_ce_id ASC"
   else  
   Sql = "select * from overtime where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,bonbu,saupbu,team,org_name,work_date,mg_ce_id ASC"
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
    <td colspan="16" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;야·특근 현황(수당)--<%=view_condi%>(<%=from_date%>∼<%=to_date%>)&nbsp;</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">본부</div></td>
    <td><div align="center" class="style1">사업부</div></td>
    <td><div align="center" class="style1">팀</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">상주처</div></td>
    <td><div align="center" class="style1">구분</div></td>
    <td><div align="center" class="style1">작업일시</div></td>
    <td><div align="center" class="style1">고객사 명</div></td>
    <td><div align="center" class="style1">지점명</div></td>
    <td><div align="center" class="style1">작업자</div></td>
    <td><div align="center" class="style1">전자결재No.</div></td>
    <td><div align="center" class="style1">금액</div></td>
    <td><div align="center" class="style1">AS No.</div></td>
    <td><div align="center" class="style1">작업내용</div></td>
  </tr>
    <%
		do until rs.eof 
		
		emp_no = rs("mg_ce_id")
		Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
        Set rs_emp = DbConn.Execute(SQL)
		if not Rs_emp.eof then
               emp_company = rs_emp("emp_company")
			   emp_name = rs_emp("emp_name")
			   emp_end_date = rs_emp("emp_end_date")
		end if
		rs_emp.close()
		
		if isNull(emp_end_date) or emp_end_date = "1900-01-01" or emp_end_date = "0000-00-00" then
			   emp_end = ""
		   else 
			   emp_end = "퇴직"
		end if

	%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="left" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="left" class="style1"><%=rs("bonbu")%></div></td>
    <td width="145"><div align="left" class="style1"><%=rs("saupbu")%></div></td>
    <td width="145"><div align="left" class="style1"><%=rs("team")%></div></td>
    <td width="145"><div align="left" class="style1"><%=rs("org_name")%></div></td>
    <td width="145"><div align="left" class="style1"><%=rs("reside_place")%></div></td>
    <td width="145"><div align="left" class="style1"><%=rs("cost_detail")%></div></td>
    <td width="150"><div align="left" class="style1"><%=rs("work_date")%>&nbsp;<%=mid(rs("from_time"),1,2)%>:<%=mid(rs("from_time"),3,2)%>∼<%=mid(rs("to_time"),1,2)%>:<%=mid(rs("to_time"),3,2)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("dept")%></div></td>
    <td width="150"><div align="left" class="style1"><%=emp_name%>(<%=rs("mg_ce_id")%>)<%=emp_end%></div></td>
    <td width="145"><div align="left" class="style1">연장/휴일-<%=rs("sign_no")%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("overtime_amt"),0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("acpt_no")%></div></td>
    <td width="500"><div align="left" class="style1"><%=rs("work_gubun")%>-<%=rs("work_memo")%></div></td>
  </tr>
	<%
	Rs.MoveNext()
	loop
	%>
  <tr valign="middle" class="style11">
    <th colspan="9"><div align="center" class="style1">합 계</div></th>
    <th colspan="2"><div align="center" class="style1"><%=formatnumber(overtime_count,0)%>&nbsp;건</div></th>
    <th colspan="2"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;원</div></th>
    <th colspan="2"><div align="left" class="style1">&nbsp;</div></th>
  </tr>   
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
