<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim s_tab(12,5)
dim emp_tab(12,5)
dim p_tab(12,5)
dim b_tab(12,3)
dim com_tab
com_tab = array("총괄","케이원정보통신","휴디스","케이네트웍스","코리아디엔씨")
dim year_tab(3,2)
dim date_tab(12)

view_year = request("view_year")
view_id = request("view_id")

cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
date_tab(12) = cstr(view_year) + "-12-31"
date_tab(11) = cstr(view_year) + "-11-30"
date_tab(10) = cstr(view_year) + "-10-31"
date_tab(9) = cstr(view_year) + "-09-30"
date_tab(8) = cstr(view_year) + "-08-31"
date_tab(7) = cstr(view_year) + "-07-31"
date_tab(6) = cstr(view_year) + "-06-30"
date_tab(5) = cstr(view_year) + "-05-31"
date_tab(4) = cstr(view_year) + "-04-30"
date_tab(3) = cstr(view_year) + "-03-31"
date_tab(2) = cstr(view_year) + "-02-29"
date_tab(1) = cstr(view_year) + "-01-31"

for i = 12 to 2 step -1
	date_tab(i-1) = dateadd("m",-1,date_tab(i))
next

for i = 0 to 12
	for j = 0 to 5
		s_tab(i,j) = 0
		emp_tab(i,j) = 0
	next
next
for i = 0 to 12
	for j = 0 to 3
		b_tab(i,j) = 0
	next
next

' 비용
sql = "select emp_company,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost where cost_year ='"&view_year&"' group by emp_company"
Rs.Open Sql, Dbconn, 1
	
do until rs.eof	
	for j = 1 to 4
		if com_tab(j) = rs("emp_company") then
			for i = 1 to 12
				if i < 10 then
					k = "0" + cstr(i)
				  else
				  	k = cstr(i)
				end if	
				cost = "cost_amt_" + cstr(k)
				s_tab(i,j) = cdbl(rs(cost))
			next
			exit for
		end if
	next
	rs.movenext()
loop
rs.close()

for i = 1 to 12
	s_tab(i,0) = s_tab(i,1) + s_tab(i,2) + s_tab(i,3) + s_tab(i,4) + s_tab(i,5)
next

' 인원수
for i = 1 to 12	
	emp_month = cstr(view_year) + right(("0" + cstr(i)),2)
	sql = "select emp_company,count(*) as emp_cnt from emp_master_month where (emp_month = '"&emp_month&"') and (emp_no > '100000' and emp_no < '199999') and (emp_in_date < '"&date_tab(i)&"') and (emp_end_date > '"&date_tab(i)&"' or emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date = '') group by emp_company"
	Rs.Open Sql, Dbconn, 1
		
	do until rs.eof	
		for j = 1 to 4
			if com_tab(j) = rs("emp_company") then
				emp_tab(i,j) = cdbl(rs("emp_cnt"))
				exit for
			end if
		next
		rs.movenext()
	loop
	rs.close()
next

for i = 1 to 12
	emp_tab(i,0) = emp_tab(i,1) + emp_tab(i,2) + emp_tab(i,3) + emp_tab(i,4) + emp_tab(i,5)
next

' 급여
for i = 1 to 12	
	emp_month = cstr(view_year) + right(("0" + cstr(i)),2)
	sql = "select pmg_company,sum(pmg_give_total) as pay_sum from pay_month_give where (pmg_yymm = '"&emp_month&"') group by pmg_company"
	Rs.Open Sql, Dbconn, 1
		
	do until rs.eof	
		for j = 1 to 4
			if com_tab(j) = rs("pmg_company") then
				p_tab(i,j) = cdbl(rs("pay_sum"))
				exit for
			end if
		next
		rs.movenext()
	loop
	rs.close()
next

for i = 1 to 12
	p_tab(i,0) = p_tab(i,1) + p_tab(i,2) + p_tab(i,3) + p_tab(i,4) + p_tab(i,5)
next

' 비용
sql = "select cost_year,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where (cost_year = '"&view_year&"') and (cost_center <> '회사간거래')"
rs.Open sql, Dbconn, 1
do until rs.eof
	for i = 1 to 12
		b_tab(i,2) = b_tab(i,2) + cdbl(rs(i))
	next
	rs.movenext()
loop
rs.close()
' 매출 집계
'sql = "select sales_month,sum(sales_amt) as cost from saupbu_sales where substring(sales_month,1,4) = '"&view_year&"' group by sales_month"
sql = "select substring(sales_date,1,7) as sales_month,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&cost_year&"' group by substring(sales_date,1,7)"
rs.Open sql, Dbconn, 1
do until rs.eof
	i = int(mid(rs("sales_month"),5,2))
	b_tab(i,1) = b_tab(i,1) + cdbl(rs("cost"))
	rs.movenext()
loop
rs.close()
' 손익계산
for i = 1 to 12
	if b_tab(i,2) = 0 then
		b_tab(i,1) = 0
	end if
	b_tab(i,3) = b_tab(i,1) - b_tab(i,2)
next

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>임원 정보 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<script type="text/javascript" src="/java/jquery.min.js"></script>
		<script type="text/javascript" src="/java/highcharts.js"></script>
		<script type="text/javascript" src="/java/modules/exporting.js"></script>
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/graph_line_year_max.js"></script>
	</head>
	<body>
		<div id="wrap">			
   			 <div id="container">
				<form action="" method="post" name="frm">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>
					<% if view_id = "1" then	%>
                        <div id="graph_view" style="width: 1200px; height: 600px; margin: 0 auto"></div>
					<% end if	%>
					<% if view_id = "2" then	%>
                        <div id="graph_view2" style="width: 1200px; height: 600px; margin: 0 auto"></div>
					<% end if	%>
					<% if view_id = "3" then	%>
                        <div id="graph_view3" style="width: 1200px; height: 600px; margin: 0 auto"></div>
					<% end if	%>
					<% if view_id = "4" then	%>
                        <div id="graph_view4" style="width: 1200px; height: 600px; margin: 0 auto"></div>
					<% end if	%>
					<% 
                       for i = 1 to 12
	                        for j = 0 to 5 
					%>
	                        <input name="s_tab<%=i%><%=j%>" type="hidden" value="<%=round(s_tab(i,j)/1000000)%>">
	                        <input name="emp_tab<%=i%><%=j%>" type="hidden" value="<%=emp_tab(i,j)%>">
	                        <input name="p_tab<%=i%><%=j%>" type="hidden" value="<%=round(p_tab(i,j)/1000000)%>">
                    <%
                            next
                        next
                    %>
					<% 
                       for i = 1 to 12
	                        for j = 1 to 3 
					%>
	                        <input name="b_tab<%=i%><%=j%>" type="hidden" value="<%=round(b_tab(i,j)/1000000)%>">
                    <%
                            next
                        next
                    %>
                   </td>
                  </tr>
                </table>
	            <input name="view_year" type="hidden" value="<%=view_year%>">				
	            <input name="view_id" type="hidden" value="<%=view_id%>">				
				</form>
			</div>
        </div>
    </body>
</html>

