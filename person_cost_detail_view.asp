<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
emp_no     = Request("emp_no")
cost_id    = Request("cost_id")
cost_yymm  = Request("cost_yymm")
cost_year  = cstr(mid(cost_yymm,1,4))
cost_month = cstr(mid(cost_yymm,5,2))
from_date  = cstr(cost_year) + "-" + cstr(cost_month) + "-01"
end_date   = datevalue(from_date)
end_date   = dateadd("m",1,from_date)
to_date    = cstr(dateadd("d",-1,end_date))

'Response.write from_date&" "&to_date

if cost_id = "야특근" then
	sql = "select org_name,work_date as slip_date,user_name,user_grade,work_item as slip_memo,overtime_amt as cost, work_gubun as cost_detail FROM overtime where (cancel_yn = 'N') and  (work_date >= '"&from_date&"' and work_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' order by org_name,user_name, work_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "일반경비" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,customer as slip_memo,cost,concat(account,' ',account_item) as cost_detail FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and emp_no = '"&emp_no&"' and slip_gubun = '비용' order by org_name,emp_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "대중교통" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,concat(company,' ',run_memo) as slip_memo,fare as cost, transit as cost_detail FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' and car_owner = '"&cost_id&"' order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "주행거리" then
	sql = "select org_name as org_name,run_date as slip_date,user_name,user_grade,concat(start_company,' -> ',end_company) as slip_memo,far as cost, concat(car_owner,' ',car_no,' ',oil_kind) as cost_detail FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' and car_owner = '개인' order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "주유비" then
	sql = "select org_name as org_name,run_date as slip_date,user_name,user_grade,concat(start_company,' -> ',end_company) as slip_memo,oil_price as cost, concat(car_owner,' ',car_no,' ',oil_kind) as cost_detail FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' and transit_cost.car_owner = '회사' order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "주차료" then
	sql = "select org_name as org_name,run_date as slip_date,user_name,user_grade,concat(start_company,' -> ',end_company) as slip_memo,parking as cost, concat(car_owner,' ',car_no,' ',oil_kind) as cost_detail FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' and parking > 0 order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "통행료" then
	sql = "select org_name as org_name,run_date as slip_date,user_name,user_grade,concat(start_company,' -> ',end_company) as slip_memo,toll as cost, concat(car_owner,' ',car_no,' ',oil_kind) as cost_detail FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' and toll > 0 order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "차량수리비" then
	sql = "select org_name as org_name,run_date as slip_date,user_name,user_grade,concat(start_company,' -> ',end_company) as slip_memo,repair_cost as cost, concat(car_owner,' ',car_no,' ',oil_kind) as cost_detail FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and mg_ce_id = '"&emp_no&"' and car_owner = '회사' order by org_name,user_name, transit_cost.run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "주유카드" then
	sql = "select org_name,card_slip.slip_date,emp_name as user_name,emp_grade as user_grade,customer as slip_memo,price as cost,concat(card_slip.account,' ',card_slip.account_item) as cost_detail FROM card_slip where card_type like '%주유%' and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and emp_no = '"&emp_no&"' order by org_name,emp_name, slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "법인카드" then
	sql = "select org_name,card_slip.slip_date,emp_name as user_name,emp_grade as user_grade,customer as slip_memo,cost as cost,concat(card_slip.account,' ',card_slip.account_item) as cost_detail FROM card_slip where card_type not like '%주유%' and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and emp_no = '"&emp_no&"' order by org_name,emp_name, slip_date asc"
	rs.Open sql, Dbconn, 1
end if
'Response.write sql & "<br>"

title_line = "개인별 비용 사용 현황"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>년월 : </strong><%=cost_year%>년<%=cost_month%>월&nbsp;
							<strong>비용구분 : </strong><%=cost_id%>&nbsp;<%=cost_detail%>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
							<col width="12%" >
							<col width="12%" >
							<col width="20%" >
							<col width="25%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">조직</th>
								<th scope="col">사용자</th>
								<th scope="col">비용일자</th>
								<th scope="col">비용구분</th>
								<th scope="col">사용내역</th>
								<th scope="col">사용금액</th>
							</tr>
						</thead>
						<tbody>
         					<% 
							cost_cnt = 0
							cost_sum = 0
							do until rs.eof
								if rs("cost") <> "0" then
									cost_sum = cost_sum + clng(rs("cost"))
									cost_cnt = cost_cnt + 1
									user_grade_view = rs("user_grade")
									slip_memo_view = rs("slip_memo")
									cost_detail = rs("cost_detail")
							%>
							<tr>
								<td class="first"><%=cost_cnt%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("user_name")%>&nbsp;<%=user_grade_view%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=cost_detail%></td>
								<td><%=slip_memo_view%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							</tr>
							<%
								end if
								rs.movenext()
							loop
							rs.close()
							%>
							<tr>
								<th colspan="6" class="first">합계</th>
								<th class="right"><%=formatnumber(cost_sum,0)%></th>
							</tr>
						</tbody>
					</table>
				</div>				        				
	</form>
	</body>
</html>

