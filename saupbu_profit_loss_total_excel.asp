<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim year_tab(5)
dim sum_amt(20,3,13)
dim saupbu_tab(20)

cost_year=Request.form("cost_year")

cost_year=Request("cost_year")

if cost_year = "" then
	cost_year = mid(cstr(now()),1,4)
	base_year = cost_year
	view_sw = "0"
end If

be_year = int(cost_year) - 1
for i = 1 to 5
	year_tab(i) = int(cost_year) - i + 1
next

for i = 1 to 20
	saupbu_tab(i) = ""
next

for i = 1 to 20
	for j = 1 to 3
		for k = 1 to 13
			sum_amt(i,j,k) = 0
		next
	next
next

' 영업조직 발췌
sql = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
rs.Open sql, Dbconn, 1
i = 0
do until rs.eof
	i = i + 1
	saupbu_tab(i) = rs("saupbu")
	rs.movenext()
loop
rs.close()

'---------------------------------------------------------------------------------------------------------------
'// 2017-09-15 회계재무 팀만 기타사업부,회사간거래 조회 가능하게 수정
'---------------------------------------------------------------------------------------------------------------
If team="회계재무" Then
	i = i + 1 
	saupbu_tab(i) = "기타사업부"
	i = i + 1 
	saupbu_tab(i) = "회사간거래"

	' 회사간거래
	sql = "select cost_center,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where cost_year = '"&cost_year&"' and (cost_center = '회사간거래') group by cost_center"
	rs.Open sql, Dbconn, 1
	do until rs.eof
		for k = 1 to 12
			sum_amt(i,2,k) = sum_amt(i,2,k) + cdbl(rs(k))
		next
		rs.movenext()
	loop
	rs.close()
End If
'---------------------------------------------------------------------------------------------------------------

' 매출 집계
sql = "select substring(sales_date,1,7) as sales_month,saupbu,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&cost_year&"' group by substring(sales_date,1,7), saupbu"
rs.Open sql, Dbconn, 1
do until rs.eof
	for i = 1 to 20
		if saupbu_tab(i) = rs("saupbu") then
			j = 1
			k = int(mid(rs("sales_month"),6,2))
			sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs("cost"))
			exit for
		end if
	next			
	rs.movenext()
loop
rs.close()

' 비용 집계
sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' group by saupbu"
rs.Open sql, Dbconn, 1
do until rs.eof
	for i = 1 to 20
		if saupbu_tab(i) = rs("saupbu") then
			j = 2
			for k = 1 to 12
				sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs(k))
			next
			exit for
		end if
	next			
	rs.movenext()
loop
rs.close()

' 비용 집계 (기타사업부)
sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and saupbu = '' group by saupbu"
rs.Open sql, Dbconn, 1
do until rs.eof
	for i = 1 to 20
		if saupbu_tab(i) = "기타사업부" then
			j = 2
			for k = 1 to 12
				sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs(k))
			next
			exit for
		end if 
	next			
	rs.movenext()
loop
rs.close()

' 비용이 없으면 매출도 표기 하지 않음
'for i = 1 to 20
'	if saupbu_tab(i) = "" then
'		exit for
'	end if
'	for k = 1 to 12
'		if sum_amt(i,2,k) = 0 then
'			sum_amt(i,1,k) = 0
'		end if
'	next
'next			

' 손익계산
for i = 1 to 20
	if saupbu_tab(i) = "" then
		exit for
	end if
	j = 3
	for k = 1 to 12
		sum_amt(i,j,k) = sum_amt(i,1,k) - sum_amt(i,2,k)
	next
next			

' 년 합계
for i = 1 to 20
	if saupbu_tab(i) = "" then
		exit for
	end if
	for j = 1 to 3
		for k = 1 to 12
			sum_amt(i,j,13) = sum_amt(i,j,13) + sum_amt(i,j,k)
		next
	next
next			

' 총계
for i = 1 to 20
	if saupbu_tab(i) = "" then
		exit for
	end if
	for j = 1 to 3
		for k = 1 to 13
			sum_amt(0,j,k) = sum_amt(0,j,k) + sum_amt(i,j,k)
		next
	next
next			

title_line = cost_year + "년" + " 사업부별 손익 총괄 현황"
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
							<col width="*" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
							  <th class="first" scope="col">사업부</th>
							  <th scope="col">구분</th>
						<% for i = 1 to 12	%>
							  <th scope="col"><%=i%>월</th>
						<% next	%>
							  <th scope="col">합계</th>
                          </tr>
						</thead>
						<tbody>
					<%
						for i = 1 to 20
							if saupbu_tab(i) = "" then
								exit for
							end if
					%>							
							<tr>
							  	<td rowspan="3" class="first"><%=saupbu_tab(i)%></td>
								<td>매출</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(i,1,k),0)%></td>
						<%	
							next	
						%>								
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">비용</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right">
								<%=formatnumber(sum_amt(i,2,k),0)%>
                                </td>
						<%	
							next	
						%>								
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">손익</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(i,3,k),0)%></td>
						<%	
							next	
						%>								
			              </tr>
					<%
						next
					%>
							<tr>
							  	<td rowspan="3" class="first" bgcolor="#CCFFFF"><strong>계</strong></td>
								<td>매출</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(0,1,k),0)%></td>
						<%	
							next	
						%>								
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">비용</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(0,2,k),0)%></td>
						<%	
							next	
						%>								
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">손익</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(0,3,k),0)%></td>
						<%	
							next	
						%>								
			              </tr>
						</tbody>
					</table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

