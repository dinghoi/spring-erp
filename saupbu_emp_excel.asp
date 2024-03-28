<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	 
cost_month=Request("cost_month")
sales_saupbu=Request("sales_saupbu")

if cost_month = "" then
	before_date = dateadd("m",-1,now())
	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
	sales_saupbu = "전체"
end If
cost_date = mid(cstr(cost_month),1,4) + "-" + mid(cstr(cost_month),5,2) + "-01"
start_date = dateadd("m",-1,cost_date)

'sql = "select * from emp_master_month where emp_month = '"&cost_month&"' and mg_saupbu = '"&sales_saupbu&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&cost_date&"') order by emp_bonbu, emp_saupbu, emp_team, emp_org_name, emp_reside_place, emp_reside_company, emp_name"

if sales_saupbu = "전체" then
	sql = "SELECT emp_master_month.*, pay_month_give.pmg_job_support, pay_month_give.pmg_give_total FROM emp_master_month INNER JOIN pay_month_give ON (emp_master_month.emp_no = pay_month_give.pmg_emp_no) AND (emp_master_month.emp_month = pay_month_give.pmg_yymm) WHERE (emp_master_month.emp_month='"&cost_month&"') and (pmg_id = '1') and (emp_master_month.cost_center <> '손익제외') order by emp_bonbu, emp_saupbu, emp_team, emp_org_name, emp_reside_place, emp_reside_company, emp_name"
  else	
	sql = "SELECT emp_master_month.*, pay_month_give.pmg_job_support, pay_month_give.pmg_give_total FROM emp_master_month INNER JOIN pay_month_give ON (emp_master_month.emp_no = pay_month_give.pmg_emp_no) AND (emp_master_month.emp_month = pay_month_give.pmg_yymm) WHERE (emp_master_month.emp_month='"&cost_month&"') and (emp_master_month.mg_saupbu = '"&sales_saupbu&"') and (pmg_id = '1') and (emp_master_month.cost_center <> '손익제외') order by emp_bonbu, emp_saupbu, emp_team, emp_org_name, emp_reside_place, emp_reside_company, emp_name"
end if

rs.Open sql, Dbconn, 1
	
title_line = cost_month + "월 " + sales_saupbu + " 인원 현황"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
					<table border="1" cellpadding="0" cellspacing="0" width="100%">
						<colgroup>
							<col width="3%" >
							<col width="*" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="10%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">상주처</th>
								<th scope="col">상주회사</th>
								<th scope="col">사번</th>
								<th scope="col">사원명</th>
								<th scope="col">직위</th>
								<th scope="col">퇴사일</th>
								<th scope="col">비용구분</th>
								<th scope="col">관리본부</th>
								<th scope="col">급여총액</th>
								<th scope="col">야특근</th>
								<th scope="col"></th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						j = 0
						team_sum = 0
						team_overtim_sum = 0
						tot_sum = 0
						tot_overtime_sum = 0
						bi_team = "first"
						do until rs.eof
							if bi_team = "first" then
								bi_team = rs("emp_team")
							end if
							if bi_team <> rs("emp_team") then
						%>
							<tr bgcolor="#FFFFCC">
								<td colspan="2" class="first">소계</td>
								<td>인원수&nbsp;&nbsp;<%=j%></td>
								<td><%=bI_team%>&nbsp;</td>
								<td colspan="8">&nbsp;</td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(team_sum,0)%>
 							<%   else	%>
								********
                            <% end if	%>
                                </td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(team_overtime_sum,0)%>
 							<%   else	%>
								********
                            <% end if	%>
                                </td>
								<td></td>
							</tr>
                        <%
								j = 0
								bi_team = rs("emp_team")								
								team_sum = 0
								team_overtime_sum = 0
							end if
							i = i + 1
							j = j + 1
'							emp_end_date = rs("emp_end_date")
'							if emp_end_date = "1900-01-01" then
'								emp_end_date = ""
'							end if
'							sql = "select pmg_give_total,pmg_job_support from pay_month_give where pmg_yymm = '"&cost_month&"' and pmg_id ='1' and pmg_emp_no ='"&rs("emp_no")&"'"
'							Set rs_etc=DbConn.Execute(Sql)
'							if rs_etc.eof or rs_etc.bof then
'								pmg_give_total = 0
'								pmg_job_support = 0
'							  else
'							  	pmg_give_total = rs_etc("pmg_give_total")
'							  	pmg_job_support = rs_etc("pmg_job_support")
'							end if
						  	pmg_give_total = rs("pmg_give_total")
						  	pmg_job_support = rs("pmg_job_support")

							team_sum = team_sum + pmg_give_total
							team_overtime_sum = team_overtime_sum + pmg_job_support
							tot_sum = tot_sum + pmg_give_total
							tot_overtime_sum = tot_overtime_sum + pmg_job_support
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("emp_bonbu")%>&nbsp;</td>
								<td><%=rs("emp_saupbu")%>&nbsp;</td>
								<td><%=rs("emp_team")%>&nbsp;</td>
								<td><%=rs("emp_reside_place")%>&nbsp;</td>
								<td><%=rs("emp_reside_company")%>&nbsp;</td>
								<td><%=rs("emp_no")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("emp_job")%></td>
								<td><%=emp_end_date%>&nbsp;</td>
								<td><%=rs("cost_center")%></td>
								<td><%=rs("mg_saupbu")%>&nbsp;</td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(pmg_give_total,0)%>
 							<%   else	%>
								********
                            <% end if	%>
                                </td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(pmg_job_support,0)%>
 							<%   else	%>
								********
                            <% end if	%>
                                </td>
								<td></td>
							</tr>
						<%
							rs.movenext()
						loop
						%>
							<tr bgcolor="#FFFFCC">
								<td colspan="2" class="first">소계</td>
								<td>인원수&nbsp;&nbsp;<%=j%></td>
								<td><%=bI_team%>&nbsp;</td>
								<td colspan="8">&nbsp;</td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(team_sum,0)%>
 							<%   else	%>
								********
                            <% end if	%>
								</td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(team_overtime_sum,0)%>
 							<%   else	%>
								********
                            <% end if	%>
								</td>
                                <td></td>
							</tr>
							<tr bgcolor="#FFE8E8">
								<td colspan="2" class="first">총계</td>
								<td>인원수&nbsp;&nbsp;<%=i%></td>
								<td>&nbsp;</td>
								<td colspan="8">&nbsp;</td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(tot_sum,0)%>
 							<%   else	%>
								********
                            <% end if	%>
								</td>
								<td class="right">
							<% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
								<%=formatnumber(tot_overtime_sum,0)%>
 							<%   else	%>
								********
                            <% end if	%>
								</td>
								<td></td>
							</tr>
						</tbody>
					</table>
		</div>				
	</div>        				
	</body>
</html>

