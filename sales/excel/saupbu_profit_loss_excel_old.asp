<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim sum_amt(10)
dim tot_amt(10)
dim detail_tab(30)
dim cost_amt(30,10)
dim cost_tab

cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

cost_year=Request("cost_year")
cost_mm=Request("cost_mm")
sales_saupbu=Request("sales_saupbu")
cost_month = cstr(cost_year) + cstr(cost_mm)

if cost_mm = "01" then
	before_year = cstr(int(cost_year) - 1)
	before_mm = "12"
  else
  	before_year = cost_year
	before_mm = right("0" + cstr(int(cost_mm) - 1),2)
end if

before_month = cstr(before_year) + cstr(before_mm)
c_month = cstr(cost_year) + "-" + cstr(cost_mm)
b_month = cstr(before_year) + "-" + cstr(before_mm)

if sales_saupbu = "전체" then
	condi_sql = ""
  else
  	condi_sql = " and saupbu ='"&sales_saupbu&"'"
end if
if sales_saupbu = "기타사업부" then
  	condi_sql = " and (saupbu ='' or saupbu = '기타사업부')"
end if	

for i = 0 to 10
	sum_amt(i) = 0
	tot_amt(i) = 0
next

sql = "select sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&b_month&"'"&condi_sql
Set rs_sum = Dbconn.Execute (sql)	
if isnull(rs_sum(0)) then
	before_sales_amt = 0 
  else
	before_sales_amt = CCur(rs_sum(0)) 
end if

sql = "select sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&c_month&"'"&condi_sql
Set rs_sum = Dbconn.Execute (sql)	
if isnull(rs_sum(0)) then
	curr_sales_amt = 0 
  else
	curr_sales_amt = CCur(rs_sum(0)) 
end if

title_line = cost_year + "년" + cost_mm + "월 " + sales_saupbu + " 사업부별 손익 현황"
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
							<col width="5%" >
							<col width="*" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
							<col width="1%" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">비용항목</th>
							  <th rowspan="2" scope="col">세부내역</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 월&nbsp;(<%=before_year%>년<%=before_mm%>월)</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">당 월&nbsp;(<%=cost_year%>년<%=cost_mm%>월)</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">증감</th>
							  <th rowspan="2" scope="col"></th>
						  </tr>
							<tr>
							  <th scope="col" style="border-left:1px solid #e3e3e3;">상주직접비</th>
							  <th scope="col">직접비</th>
							  <th scope="col">전사공통비</th>
							  <th scope="col">부문공통비</th>
							  <th scope="col">계</th>
							  <th scope="col">상주직접비</th>
							  <th scope="col">직접비</th>
							  <th scope="col">전사공통비</th>
							  <th scope="col">부문공통비</th>
							  <th scope="col">계</th>
							  <th scope="col">금액</th>
							  <th scope="col">율</th>
                          </tr>
						</thead>
						<tbody>
						<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>매출계</strong></td>
							  <td colspan="5" scope="col" class="right"><%=formatnumber(before_sales_amt,0)%></td>
							  <td colspan="5" scope="col" class="right"><%=formatnumber(curr_sales_amt,0)%></td>
						<%
						   incr_amt = curr_sales_amt - before_sales_amt
						   if before_sales_amt = 0 and curr_sales_amt = 0 then
						   		incr_per = 0 
							  elseif before_sales_amt = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / before_sales_amt * 100
						   end if
						%>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
					<%
					for jj = 0 to 10
						rec_cnt = 0

						for i = 1 to 30
							detail_tab(i) = ""
							for j = 1 to 10
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
							sql = "select cost_detail from saupbu_profit_loss where (cost_year ='"&cost_year&"' or cost_year ='"&before_year&"') and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rs("cost_detail")
								rs.movenext()
							loop
							rs.close()
						end if
						if rec_cnt <> 0 then
' 전월 금액 SUM
							sql = "select cost_center,cost_detail,sum(cost_amt_"&before_mm&") as cost from saupbu_profit_loss where cost_year ='"&before_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_center,cost_detail order by cost_center, cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								for i = 1 to 30
									if rs("cost_detail") = detail_tab(i) then
										select case rs("cost_center")
											case "상주직접비"
												j = 1
											case "직접비"
												j = 2
											case "전사공통비"
												j = 3
											case "부문공통비"
												j = 4									
										end select
										cost_amt(i,j) = cost_amt(i,j) + ccur(rs("cost"))
										cost_amt(i,5) = cost_amt(i,5) + ccur(rs("cost"))
										sum_amt(j) = sum_amt(j) + ccur(rs("cost"))
										sum_amt(5) = sum_amt(5) + ccur(rs("cost"))
										tot_amt(j) = tot_amt(j) + ccur(rs("cost"))
										tot_amt(5) = tot_amt(5) + ccur(rs("cost"))
										exit for
									end if
								next								
								rs.movenext()
							loop
							rs.close()
' 당월 금액 SUM
							sql = "select cost_center,cost_detail,sum(cost_amt_"&cost_mm&") as cost from saupbu_profit_loss where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_center,cost_detail order by cost_center, cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								for i = 1 to 30
									if rs("cost_detail") = detail_tab(i) then
										select case rs("cost_center")
											case "상주직접비"
												j = 6
											case "직접비"
												j = 7
											case "전사공통비"
												j = 8
											case "부문공통비"
												j = 9									
										end select
										cost_amt(i,j) = cost_amt(i,j) + ccur(rs("cost"))
										cost_amt(i,10) = cost_amt(i,10) + ccur(rs("cost"))
										sum_amt(j) = sum_amt(j) + ccur(rs("cost"))
										sum_amt(10) = sum_amt(10) + ccur(rs("cost"))
										tot_amt(j) = tot_amt(j) + ccur(rs("cost"))
										tot_amt(10) = tot_amt(10) + ccur(rs("cost"))
										exit for
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
						<% for j = 1 to 10	%>
								<td class="right"><%=formatnumber(cost_amt(1,j),0)%></td>
						<% next	%>
						<%
						   incr_amt = cost_amt(1,10) - cost_amt(1,5)
						   if cost_amt(1,5) = 0 and cost_amt(1,10) = 0 then
						   		incr_per = 0 
							  elseif cost_amt(1,5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / cost_amt(1,5) * 100
						   end if
						%>
								<td class="right"><%=formatnumber(incr_amt,0)%></td>
				        <td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right">&nbsp;</td>
							</tr>
						<% for i = 2 to rec_cnt	%>
                        	<tr>
								<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
						<%   for j = 1 to 10	%>
								<td class="right"><%=formatnumber(cost_amt(i,j),0)%></td>
						<%   next	%>
						<%
						   incr_amt = cost_amt(i,10) - cost_amt(i,5)
						   if cost_amt(i,5) = 0 and cost_amt(i,10) = 0 then
						   		incr_per = 0 
							  elseif cost_amt(i,5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / cost_amt(i,5) * 100
						   end if
						%>
					      <td class="right"><%=formatnumber(incr_amt,0)%></td>
								<td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right">&nbsp;</td>
							</tr>
						<% next	%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
						<% for j = 1 to 10	%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(j),0)%></td>
						<% next	%>
						<%
						   incr_amt = sum_amt(10) - sum_amt(5)
						   if sum_amt(5) = 0 and sum_amt(10) = 0 then
						   		incr_per = 0 
							  elseif sum_amt(5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / sum_amt(5) * 100
						   end if
						%>
					      <td class="right" bgcolor="#EEFFFF"><%=formatnumber(incr_amt,0)%></td>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right" bgcolor="#EEFFFF">&nbsp;</td>
							</tr>
					<%
						end if
					next
					%>
					<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>비용합계</strong></td>
						<% for j = 1 to 10	%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(j),0)%></td>
						<% next	%>
						<%
						   incr_amt = tot_amt(10) - tot_amt(5)
						   if tot_amt(5) = 0 and tot_amt(10) = 0 then
						   		incr_per = 0 
							  elseif tot_amt(5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / tot_amt(5) * 100
						   end if
						%>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
						<tr bgcolor="#FFDFDF">
							  <td colspan="2" class="first" scope="col"><strong>손익</strong></td>
						<%
						   	be_profit_loss = before_sales_amt - tot_amt(5)	
						   	curr_profit_loss = curr_sales_amt - tot_amt(10)	
						   	incr_amt = curr_profit_loss - be_profit_loss
						   	if be_profit_loss = 0 and curr_profit_loss = 0 then
						   		incr_per = 0 
							  elseif be_profit_loss = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / be_profit_loss * 100
						   	end if
							if be_profit_loss < 0 then
								incr_per = incr_per * -1
							end if 
						%>
							  <td scope="col" colspan="5" class="right"><%=formatnumber(be_profit_loss,0)%></td>
							  <td scope="col" colspan="5" class="right"><%=formatnumber(curr_profit_loss,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
						</tbody>
					</table>
				
		</div>				
	</div>        				
	</body>
</html>

