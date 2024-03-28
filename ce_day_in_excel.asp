<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

dim in_cnt_tab(31)
dim in_tot_tab(31)
dim in_date_tab(31)

from_date = Request("from_date")
to_date = Request("to_date")
team = "전체"

savefilename = "CE별 일자별 입고현황" + to_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

in_cnt_tab(0) = 0
in_tot_tab(0) = 0
for i = 0 to 30
	in_date_tab(i+1) = mid(cstr(dateadd("d",i,from_date)),1,10)
	in_cnt_tab(i+1) = 0
	in_tot_tab(i+1) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if  team = "전체" then
	sql = "select memb.user_id,memb.team,memb.user_name,memb.reside from as_acpt inner join memb on as_acpt.mg_ce_id = memb.user_id "
	sql = sql + " Where (as_acpt.mg_group='"+mg_group+"')"
	sql = sql + " and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"')"
	sql = sql + " GROUP BY memb.user_id,memb.team,memb.user_name,memb.reside Order By memb.team, memb.user_name Asc"
  else
	sql = "select memb.user_id,memb.team,memb.user_name,memb.reside from as_acpt inner join memb on as_acpt.mg_ce_id = memb.user_id "
	sql = sql + " Where (as_acpt.mg_group='"+mg_group+"') and (memb.team='"+team+"')"
	sql = sql + " and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"')"
	sql = sql + " GROUP BY memb.user_id,memb.team,memb.user_name,memb.reside Order By memb.user_name Asc"
end if
Rs.Open Sql, Dbconn, 1
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
</head>

<body>
<table width="1200" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="30" bgcolor="#6699CC">&nbsp;<span class="style14BW">CE별 일자별 입고 현황</span></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellspacing="3">
      <tr>
        <td>
          <table width="1200" border="0" cellspacing="0" cellpadding="0">
            <tr valign="middle" class="style6">
              <td width="100" height="25" bgcolor="#CCCCCC"><div align="center" class="style6">기준일</div></td>
              <td height="25">&nbsp;<%=from_date%> - <%=to_date%></td>
              </tr>
          </table>
        </td>
      </tr>
              </table></td>
            </tr>
            <tr>
              <td><table width="100%" border="1" cellspacing="0" cellpadding="0">
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td width="100" rowspan="2" bgcolor="#FFFF99"><div align="center" class="style12">소속</div></td>
                  <td width="90" rowspan="2" bgcolor="#FFFF99"><div align="center">CE</div></td>
                  <td width="50" rowspan="2" bgcolor="#FFFF99"><div align="center">상주</div></td>
                  <td height="20" colspan="32"><div align="center">일자별</div></td>
                </tr>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td width="30" height="20"><div align="center" class="style12">계</div></td>
				<%
				for i = 1 to 31	
				%>
                  <td width="30" height="20"><div align="center"><%=right(in_date_tab(i),2)%></div></td>
				<%
				next
				%>
                </tr>
                <% 
				do until rs.eof 
		' 월간 미처리 입고
					sql = "select count(*) as in_cnt, in_date from as_acpt "
					sql = sql + "WHERE (mg_ce_id='"+rs("user_id")+"') and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"') GROUP BY in_date Order By in_date Asc"
					Rs_in.Open Sql, Dbconn, 1
					do until rs_in.eof
						in_cnt = clng(rs_in("in_cnt"))
						for j = 1 to 31
							if cstr(rs_in("in_date")) = cstr(in_date_tab(j)) then
								in_cnt_tab(j) = in_cnt_tab(j) + in_cnt
								in_cnt_tab(0) = in_cnt_tab(0) + in_cnt				
								in_tot_tab(j) = in_tot_tab(j) + in_cnt
								in_tot_tab(0) = in_tot_tab(0) + in_cnt				
								exit for
							end if 
						next
						rs_in.movenext()
					loop
					rs_in.close()

					if rs("reside") = "0" then
						reside = "."
					  else
						reside = "상주"
					end if
				%>
                <tr class="style12">
                  <td width="100" height="20"><div align="center"><%=rs("team")%></div></td>
                  <td width="90" height="20"><div align="center"><%=rs("user_name")%></div></td>
                  <td width="50" height="20"><div align="center"><%=reside%></div></td>
                  <td width="30" bgcolor="#CCFFCC"><div align="center"><%=in_cnt_tab(0)%></div></td>
				<% for j = 1 to 31 %>
                  <td width="30"><div align="center"><%=in_cnt_tab(j)%></div></td>
 				<%	next %>
                </tr>
                <%
					for i = 0 to 31
						in_cnt_tab(i) = 0
					next
					rs.movenext()
				loop
				rs.close()
				%>
                <tr valign="middle" bgcolor="#FFFFFF" class="style12">
                  <td height="20" colspan="3" bgcolor="#CCCCCC"><div align="center"><strong>계</strong></div></td>
                  <td width="30" bgcolor="#CCFFCC"><div align="center"><%=in_tot_tab(0)%></div></td>
				<% for j = 1 to 31 %>
                  <td width="30"><div align="center"><%=in_tot_tab(j)%></div></td>
 				<%	next %>
                </tr>
              </table></td>
            </tr>
          </table>
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
