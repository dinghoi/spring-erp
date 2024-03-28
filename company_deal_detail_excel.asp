<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
cost_month = request("cost_month")
cost_year = mid(cost_month,1,4)
cost_mm = mid(cost_mm,5,2)

from_date = mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

sql = "select * FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and cost_center = '회사간거래' order by slip_date asc"
rs.Open sql, Dbconn, 1

title_line = "회사간 거래 세부 내역"
title_line = cost_year + "년" + cost_mm + "월 " + " 회사간 거래 세부 내역"
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
		<title>영업 관리 시스템</title>
	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">일자</th>
								<th scope="col">비용구분</th>
								<th scope="col">세부비용</th>
								<th scope="col">고객사</th>
								<th scope="col">거래처</th>
								<th scope="col">사용내역</th>
								<th scope="col">사용금액</th>
							</tr>
						</thead>
						<tbody>
         					<% 
							cost_cnt = 0
							cost_sum = 0
							i = 0
							do until rs.eof
								i = i + 1
								if rs("cost") <> "0" then
									cost_sum = cost_sum + clng(rs("cost"))
									cost_cnt = cost_cnt + 1
							%>
							<tr>
								<td class="first"><%=cost_cnt%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("company")%></td>
								<td class="left"><%=rs("customer")%></td>
								<td class="left"><%=rs("slip_memo")%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							</tr>
							<%
								end if
								rs.movenext()
							loop
							rs.close()
							%>
						</tbody>
					</table>
				</div>				        				
		</div>				        				
	</body>
</html>

