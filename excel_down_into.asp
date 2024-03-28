<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs

replace_sw=Request("replace_sw")
company=Request("company")

curr_date = datevalue(mid(cstr(now()),1,10))
savefilename = "입고진행관리" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_last = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

replace_sql = ""
if replace_sw <> "전체" then
	if replace_sw = "대체" then
		replace_sql = " and (in_replace = '대체')"
	  else
	  	replace_sql = " and (in_replace <> '대체')"
	end if
end if
company_sql = ""
if company <> "전체" then
	company_sql = " and (company = '"+company+"')"
end if
where_sql = " WHERE (mg_group = '" + mg_group + "') and (as_process = '입고' or as_process = '대체입고') "
order_sql = " ORDER BY acpt_date ASC"
condi_Sql = " and (mg_ce_id = '" + c_id + "')"

if c_grade = "0" or ( c_grade = "1" and c_belong = "수도권지사" ) then
	condi_Sql = " "
end if	
if ( c_grade = "1" and c_belong <> "수도권지사" ) then
	condi_Sql = " and (belong = '"+c_belong+"' or mg_ce_id = '"+c_id+"')"
end if	
if c_grade = "2" then
	condi_Sql = " and (reside_place = '"+reside_place+"' or mg_ce_id = '"+c_id+"')"
end if
if c_grade = "3"  and c_belong <> "수도권지사" then
	condi_Sql = " and (belong = '"+c_belong+"' or mg_ce_id = '"+c_id+"')"
end if
if c_grade = "3"  and c_belong = "수도권지사" then
	Sql = " and (mg_ce_id = '"+c_id+"')"
end if
sql = "select * from as_acpt " + where_sql + condi_sql + replace_sql + company_sql + order_sql
Rs.Open Sql, Dbconn, 1
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 10px}
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=now()%> &nbsp;입고 진행 현황</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">경과</div></td>
    <td><div align="center" class="style1">접수일자</div></td>
    <td><div align="center" class="style1">고객명</div></td>
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">조직명</div></td>
    <td><div align="center" class="style1">담당CE</div></td>
    <td><div align="center" class="style1">제조사</div></td>
    <td><div align="center" class="style1">입고장비</div></td>
    <td><div align="center" class="style1">입고처</div></td>
    <td><div align="center" class="style1">대체</div></td>
    <td><div align="center" class="style1">최종처리</div></td>
    <td><div align="center" class="style1">
        <div align="left">입고 세부내역 </div>
    </div></td>
  </tr>
  <%
		do until rs.eof 

'휴일 계산
			hol_d = 0
			com_date = datevalue(mid(rs("acpt_date"),1,10))
			dd = datediff("d", com_date, curr_date)
			if dd > 0 then
				a = datediff("d", com_date, curr_date)
				b = datepart("w",com_date)
				c = a + b
				d = a
				if a > 1 then
					if c > 7 then
						d = a - 2
					end if
				end if
				
		'		visit_date = rs("visit_date")
		'		act_date = com_date
			
				do until com_date > curr_date
					sql_hol = "select * from holiday where holiday = '" + cstr(com_date) + "'"
					Set rs_hol=DbConn.Execute(SQL_hol)
					if rs_hol.eof or rs_hol.bof then
						d = d
					  else 
						d = d -1
					end if
					com_date = dateadd("d",1,com_date)
					rs_hol.close()
				loop
	
				if d > 6 then
					hol_d = int(d/7) * 2
				end if
	'			if d > 2 then
	'				d = 3
	'			end if
	'			if d = 1 then
	'				j = 5
	'			  elseif d = 2 then
	'				j = 6
	'			  else
	'				j = 7
	'			end if
				d_day = d - hol_d
			  else
		' 휴일 계산 끝
				d_day = 0
			end if

			if rs("in_replace") = "" or isnull(rs("in_replace")) then
				in_replace = "."
			  else
				in_replace = rs("in_replace")
			end if
			sql = "select * from as_into where acpt_no="&rs("acpt_no")&" order by in_seq asc"
			rs_in.Open Sql, Dbconn, 1

			sql_last = "select in_process from as_into where acpt_no="&rs("acpt_no")&" and in_seq="&"(select max(in_seq) from as_into where acpt_no="&rs("acpt_no")&")"
			Set rs_last=dbconn.execute(sql_last)
			last_process = rs_last("in_process")
			rs_last.close()
			k = 0
			do until rs_in.eof
				in_memo = cstr(rs_in("into_date")) + "/" + rs_in("in_place") + "/"  + rs_in("in_process") + "/" + rs_in("in_remark")
				k = k + 1		
				
'				if last_process <> "처리완료" then
'					last_process = rs_in("in_process")
'				end if
	%>
	<% if k = 1 then %>
  <tr valign="middle" class="style11">
    <td width="36"><div align="center" class="style1"><%=d_day%></div></td>
    <td width="182"><div align="center" class="style1"><%=rs("acpt_date")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("acpt_user")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("dept")%></div></td>
    <td width="54"><div align="center" class="style1"><%=rs("mg_ce")%></div></td>
    <td width="78"><div align="center" class="style1"><%=rs("maker")%></div></td>
    <td width="67"><div align="center" class="style1"><%=rs("as_device")%></div></td>
    <td width="67"><div align="center" class="style1"><%=rs_in("in_place")%></div></td>
    <td width="48"><div align="center" class="style1"><%=in_replace%></div></td>
    <td width="97"><div align="center" class="style1"><%=last_process%></div></td>
    <td width="197"><div align="center" class="style1"><div align="left"><%=in_memo%></div></div></td>
		<% else %>
    <td width="197"><div align="center" class="style1"><div align="left"><%=in_memo%></div></div></td>
  	<%
	   end if
				rs_in.movenext()
			loop
			rs_in.close()
	%>
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
