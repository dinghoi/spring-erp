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
com_tab = array("총괄","케이원정보통신","휴디스","케이네트웍스","에스유에이치","코리아디엔씨")
dim year_tab(3,2)
dim date_tab(12)

year_tab(3,1) = mid(dateadd("m",-1,now()),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

view_year = request.form("view_year")
if view_year = "" then
	view_year = mid(dateadd("m",-1,now()),1,4)
end if

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

for i = 0 to 12
	for j = 0 to 5
		s_tab(i,j) = 0
		emp_tab(i,j) = 0
		p_tab(i,j) = 0
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
	for j = 1 to 5
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
		for j = 1 to 5
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
		for j = 1 to 5
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
		if isnull(rs(i)) then
			b_tab(i,2) = b_tab(i,2) + 0
		  else	
			b_tab(i,2) = b_tab(i,2) + cdbl(rs(i))
		end if
	next
	rs.movenext()
loop
rs.close()
' 매출 집계
sql = "select substring(sales_date,1,7) as sales_month,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&view_year&"' group by substring(sales_date,1,7)"
rs.Open sql, Dbconn, 1
do until rs.eof
	i = int(mid(rs("sales_month"),6,2))
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
		<script type="text/javascript" src="/java/graph_line_year.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_year.value == "") {
					alert ("조회년을 입력하세요");
					return false;
				}	
				return true;
			}
			function pop_graph(view_year,view_id)
			{ 
				var popupW = 1200;
				var popupH = 600;
				var left = Math.ceil((window.screen.width - popupW)/2);
				var top = Math.ceil((window.screen.height - popupH)/2);
				window.open('ceo_total_graph.asp?view_year='+view_year+'&view_id='+view_id+'', '팝업그래프', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
   			 <div id="container">
				<form action="" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
							<label>
							&nbsp;&nbsp;<strong>조회년&nbsp;</strong> : 
                            <select name="view_year" id="view_year" style="width:150px">
                            <%	for i = 3 to 1 step -1	%>
                            	<option value="<%=year_tab(i,1)%>" <%If view_year = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                            <%	next	%>
                            </select>
							</label>
                            <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
							&nbsp;&nbsp;<strong>해당 그래프를 더블클릭하면 그래프 확대가 됩니다.</strong>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="49%" height="260px">
                        <div id="graph_view" style="width: 588px; height: 260px; margin: 0 auto" onDblClick="pop_graph(<%=view_year%>,<%="1"%>);"></div>
						<% 
                        for i = 1 to 12
	                        for j = 0 to 5
                        %>
	                        <input name="s_tab<%=i%><%=j%>" type="hidden" value="<%=round(s_tab(i,j)/1000000)%>">
                        <%
                            next
                        next
                        %>
                   	</td>
                    <td width="2%">&nbsp;</td>
                    <td width="49%" height="260px">
                        <div id="graph_view2" style="width: 588px; height: 260px; margin: 0 auto" onDblClick="pop_graph(<%=view_year%>,<%="2"%>);"></div>
						<% 
                        for i = 1 to 12
	                        for j = 0 to 5
                        %>
	                        <input name="emp_tab<%=i%><%=j%>" type="hidden" value="<%=emp_tab(i,j)%>">
                        <%
                            next
                        next
                        %>
                    </td>
                  </tr>
                  <tr>
                    <td height="5px">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="49%" height="260px">
                        <div id="graph_view3" style="width: 588px; height: 260px; margin: 0 auto" onDblClick="pop_graph(<%=view_year%>,<%="3"%>);"></div>
						<% 
                        for i = 1 to 12
	                        for j = 0 to 5
                        %>
	                        <input name="p_tab<%=i%><%=j%>" type="hidden" value="<%=round(p_tab(i,j)/1000000)%>">
                        <%
                            next
                        next
                        %>
                   	</td>
                    <td width="2%">&nbsp;</td>
                    <td width="49%" height="260px">
                        <div id="graph_view4" style="width: 588px; height: 260px; margin: 0 auto" onDblClick="pop_graph(<%=view_year%>,<%="4"%>);"></div>
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
				</form>
			</div>
        </div>
    </body>
</html>

