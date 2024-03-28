<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim sum_amt(8)
dim tot_amt(8)
dim detail_tab(30)
dim cost_amt(30,8)
dim saupbu_tab(8)
dim sales_amt(8)
dim cost_tab

cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

cost_month=Request("cost_month")

cost_year = mid(cost_month,1,4)
cost_mm = mid(cost_month,5)
c_month = cost_year + "-" + cost_mm
for i = 0 to 8
	sum_amt(i) = 0
	tot_amt(i) = 0
	sales_amt(i) = 0
next

i = 0
sql = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
rs_org.Open Sql, Dbconn, 1
do until rs_org.eof
	i = i + 1
	saupbu_tab(i) = rs_org(0)
	rs_org.movenext()
loop
rs_org.close()						
i = i + 1
saupbu_tab(i) = ""
i = i + 1
saupbu_tab(i) = "소계"

sql = "select saupbu,sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&c_month&"' group by saupbu"
rs.Open Sql, Dbconn, 1
do until rs.eof
	for i = 1 to 7
		if saupbu_tab(i) = rs("saupbu") then
			sales_amt(i) = CCur(rs("sales_amt"))
			sales_amt(8) = sales_amt(8) + CCur(rs("sales_amt"))
			exit for
		end if
	next
	rs.movenext()
loop
rs.close()						

title_line = cost_year + "년" + cost_mm + "월 " + " 사업부별 손익 현황"
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
                <div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
							  <th class="first" scope="col">비용항목</th>
							  <th scope="col">세부내역</th>
						<% for i = 1 to 6	%>
							  <th scope="col"><%=saupbu_tab(i)%></th>
						<% next	%>
							  <th scope="col">사업부미지정</th>
							  <th scope="col">소계</th>
							  <th scope="col"></th>
                          </tr>
						</thead>
						<tbody>
						<tr bgcolor="#FFFFCC">
							<td colspan="2" class="first" scope="col"><strong>매출</strong></td>
					<% for i = 1 to 8	%>				
                    		<td class="right" scope="col"><%=formatnumber(sales_amt(i),0)%></td>
 					<% next	%>
                			<td scope="col" class="right">&nbsp;</td>
                         </tr>
					<%
					for jj = 0 to 10
						rec_cnt = 0

						for i = 1 to 30
							detail_tab(i) = ""
							for j = 1 to 8
								cost_amt(i,j) = 0
								sum_amt(j) = 0
							next
						next
						if cost_tab(jj) = "인건비" then
							sql = "select cost_detail from saupbu_cost_account where cost_id ='"&cost_tab(jj)&"' order by view_seq"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rs("cost_detail")
								rs.movenext()
							loop
							rs.close()
						  else
							sql = "select cost_detail from saupbu_profit_loss where (cost_year ='"&cost_year&"') and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rs("cost_detail")
								rs.movenext()
							loop
							rs.close()
						end if
						if rec_cnt <> 0 then
' 당월 금액 SUM
							sql = "select saupbu,cost_detail,sum(cost_amt_"&cost_mm&") as cost from saupbu_profit_loss where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"' group by saupbu,cost_detail order by saupbu, cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								for i = 1 to 30
									if rs("cost_detail") = detail_tab(i) then
										for j = 1 to 7
											if saupbu_tab(j) = rs("saupbu") then
												cost_amt(i,j) = cost_amt(i,j) + clng(rs("cost"))
												cost_amt(i,8) = cost_amt(i,8) + clng(rs("cost"))
												sum_amt(j) = sum_amt(j) + clng(rs("cost"))
												sum_amt(8) = sum_amt(8) + clng(rs("cost"))
												tot_amt(j) = tot_amt(j) + clng(rs("cost"))
												tot_amt(8) = tot_amt(8) + clng(rs("cost"))
												exit for
											end if
										next
									end if
								next
								rs.movenext()
							loop
							rs.close()
						%>
							<tr>
							  	<td rowspan="<%=rec_cnt + 1%>" class="first">
						<% if jj = 2 or jj = 3 then	%>
                        	  	<%=cost_tab(jj)%><br>(현금사용)
						<%   else	%>
                        	  	<%=cost_tab(jj)%>
                        <% end if	%>
                              	</td>
								<td class="left"><%=detail_tab(1)%></td>
						<% for j = 1 to 8	%>
								<td class="right"><%=formatnumber(cost_amt(1,j),0)%></td>
						<% next	%>
								<td class="right">&nbsp;</td>
						  </tr>
					  <% for i = 2 to rec_cnt	%>
                        	<tr>
								<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
						<%   for j = 1 to 8	%>
								<td class="right"><%=formatnumber(cost_amt(i,j),0)%></td>
						<%   next	%>
								<td class="right">&nbsp;</td>
							</tr>
						<% next	%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
						<% for j = 1 to 8	%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(j),0)%></td>
						<% next	%>
								<td class="right" bgcolor="#EEFFFF">&nbsp;</td>
						  </tr>
					<%
						end if
					next
					%>
					<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>비용합계</strong></td>
						<% for j = 1 to 8	%>
								<td class="right"><%=formatnumber(tot_amt(j),0)%></td>
						<% next	%>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
						<tr bgcolor="#FFDFDF">
							  <td colspan="2" bgcolor="#FFDFDF" class="first" scope="col"><strong>손익</strong></td>
						<%
						 for j = 1 to 8	
						 	cal_amt = sales_amt(j) - tot_amt(j)
						 %>
								<td class="right"><%=formatnumber(cal_amt,0)%></td>
						<%
						 next	
						 %>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
						</tbody>
					</table>
                        </DIV>
						</td>
                    </tr>
					</table>				
			</div>				
		</div>        				
	</body>
</html>

