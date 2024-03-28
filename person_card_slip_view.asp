<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
slip_month = Request("slip_month")
emp_no = Request("emp_no")

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

sql = "select * from memb where user_id = '"&emp_no&"'"
Set rs = Dbconn.Execute(sql)
if rs.eof or rs.bof then
	emp_name = "ERROR"
	user_grade = "님"
  else
  	emp_name = rs("user_name")
	user_grade = rs("user_grade")
end if
rs.close()

sql = "select * from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no ='"&emp_no&"') order by slip_date"
Rs.Open Sql, Dbconn, 1
'Response.write sql

title_line = "카드 사용 내역"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
							<strong>년월 : </strong><%=slip_month%>&nbsp;
							<strong>사용직원 : </strong><%=emp_name%>&nbsp;<%=user_grade%>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
							<col width="12%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
							<col width="9%" >
							<col width="12%" >
							<col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">순번</th>
								<th rowspan="2" scope="col">사용일</th>
								<th rowspan="2" scope="col">카드유형</th>
								<th rowspan="2" scope="col">거래처</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">사용 금액</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">사용 내역</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">합계</th>
							  <th scope="col">공급가액</th>
							  <th scope="col">부가세</th>
							  <th scope="col">계정과목</th>
							  <th scope="col">항목</th>
		                  </tr>
						</thead>
						<tbody>
         					<% 
							sum_price = 0
							sum_cost = 0
							sum_cost_vat = 0
							i = 0
							do until rs.eof
								i = i + 1
								sum_price = sum_price + rs("price")
								sum_cost = sum_cost + rs("cost")
								sum_cost_vat = sum_cost_vat + rs("cost_vat")
							%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("card_type")%></td>
								<td><%=rs("customer")%></td>
								<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("account_item")%></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
							<tr>
								<th colspan="2" class="first">합계</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th class="right"><%=formatnumber(sum_price,0)%></th>
								<th class="right"><%=formatnumber(sum_cost,0)%></th>
								<th class="right"><%=formatnumber(sum_cost_vat,0)%></th>
								<th class="right">&nbsp;</th>
								<th class="right">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>				        				
	</form>
	</body>
</html>

