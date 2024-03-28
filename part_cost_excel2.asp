<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'on Error resume next

Dim from_date
Dim to_date
Dim win_sw

cost_month=Request("cost_month")
sales_saupbu=Request("sales_saupbu")
'Response.write cost_month
'Response.write sales_saupbu

if cost_month = "" then
	before_date = dateadd("m",-1,now())
	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
	sales_saupbu = "전체"
end If

if sales_saupbu = "전체" then
	condi_sql = ""
  else
  	condi_sql = " and saupbu ='"&sales_saupbu&"'"
end if
mm = mid(cost_month,5,2)
cost_year = mid(cost_month,1,4)


	sql = "SELECT											" & chr(13) &_
	      "	company											" & chr(13) &_
	      "	, saupbu										" & chr(13) &_
	      "	, count(acpt_no) AS remote_cnt					" & chr(13) &_
	      "	, sum(as_standard_money) AS cost_amt			" & chr(13) &_
	      "	, (sum(as_standard_money)/(SELECT sum(as_standard_money) FROM AS_ACPT WHERE 1=1 AND DATE_FORMAT( acpt_date, '%Y%m') = '"&cost_month&"' AND as_process = '완료' AND length(trim(saupbu)) > 0))*100  AS charge_per						  " & chr(13) &_
	      "FROM AS_ACPT										" & chr(13) &_
	      "WHERE 1=1										" & chr(13) &_
	      "	AND DATE_FORMAT( acpt_date, '%Y%m') = '"&cost_month&"'	" & chr(13) &_
	      "	AND as_process = '완료'							" & chr(13) &_
	      "AND length(trim(saupbu)) > 0 "&condi_sql&"		" & chr(13) &_
	      "GROUP BY company, saupbu							" & chr(13) &_
	      "ORDER BY company ASC								"

rs.Open sql, Dbconn, 1

title_line = cost_year + "년" + mm + "월 " + sales_saupbu + " 부분 공통비 배분현황(변경후)"

savefilename = title_line + ".xls"


'Response.Buffer = True
'Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
'Response.CacheControl = "public"
'Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

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
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">회사</th>
								<th scope="col">사업부</th>
								<th scope="col">건수</th>
								<th scope="col">차지율(%)</th>
								<th scope="col">부분공통비</th>
							</tr>
						</thead>
						<tbody>
						<%
						remote_sum = 0
						charge_per_sum = 0
						charge_cost_sum = 0
						i = 0
						do until rs.eof
							i = i + 1
							remote_sum = cint(rs("remote_cnt")) + remote_sum
							charge_per_sum = CDbl(rs("charge_per")) + charge_per_sum
							charge_cost_sum = CLng(rs("cost_amt")) + charge_cost_sum
						%>
							<tr>
								<td class="first"><%=rs("company")%></td>
								<td><%=rs("saupbu")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("remote_cnt"),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("charge_per"),3)%>&nbsp;%&nbsp;</td>
								<td class="right"><%=formatnumber(rs("cost_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						%>
							<tr>
								<td bgcolor="#FFE8E8" class="first">총계</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(remote_sum,0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(charge_per_sum,3)%>&nbsp;%&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(charge_cost_sum,0)%>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				<br>
		</div>
	</div>
	</body>
</html>

