<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs

field_check=Request("field_check")
field_view=Request("field_view")
view_sw=Request("view_sw")
curr_date=Request("curr_date")

curr_date = mid(cstr(now()),1,10)
savefilename = curr_date + "미수금 관리 내역.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from saupbu_sales where (sales_amt <> collect_tot_amt) "

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

if view_sw = "1" then
	view_sql = " and ( collect_due_date < '"&curr_date&"' ) "
  elseif view_sw = "2" then
	view_sql = " and ( collect_due_date >= '"&curr_date&"' ) "
  else
  	view_sql = " "
end if

order_sql = " ORDER BY emp_name, company, sales_date ASC"

Sql = "SELECT count(*) FROM saupbu_sales where (sales_amt <> collect_tot_amt) " + field_sql + view_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

sql = "select sum(sales_amt) as price,sum(collect_tot_amt) as collect from saupbu_sales where (sales_amt <> collect_tot_amt) " + field_sql + view_sql
Set rs_sum = Dbconn.Execute (sql)
if isnull(rs_sum("price")) then
	tot_sales_amt = 0
	tot_collect_tot_amt = 0
  else
	tot_sales_amt = cdbl(rs_sum("price"))
	tot_collect_tot_amt = cdbl(rs_sum("collect"))
end if

sql = base_sql + field_sql + view_sql + order_sql
Rs.Open Sql, Dbconn, 1

title_line = "미수금 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">전표번호</th>
								<th scope="col">매출일자</th>
								<th scope="col">수금예정일</th>
								<th scope="col">미수금예정일</th>
								<th scope="col">거래처명</th>
								<th scope="col">영업담당</th>
								<th scope="col">매출총액</th>
								<th scope="col">수금총액</th>
								<th scope="col">잔액</th>
								<th scope="col">변동사항</th>
								<th scope="col">미수금 사유</th>
							</tr>
						</thead>
						<tbody>
						<%
    					seq = 0
						do until rs.eof
							seq = seq + 1
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=mid(rs("slip_no"),1,17)%></td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("collect_due_date")%></td>
								<td><%=rs("unpaid_due_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("collect_tot_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_amt")-rs("collect_tot_amt"),0)%></td>
								<td><%=rs("change_memo")%>&nbsp;</td>
								<td><%=rs("unpaid_memo")%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><strong><%=formatnumber(total_record,0)%>건<strong></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(tot_sales_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_collect_tot_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_sales_amt - tot_collect_tot_amt,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

