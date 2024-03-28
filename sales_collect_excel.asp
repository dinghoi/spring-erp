<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw
dim sum_tab(4,2)

win_sw = "close"

from_date=Request("from_date")
to_date=Request("to_date")
field_check=Request("field_check")
field_view=Request("field_view")
view_sw=Request("view_sw")

savefilename = from_date + "~" + to_date + " 수금현황.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select sales_collect.*, saupbu_sales.sales_date, saupbu_sales.company, saupbu_sales.sales_amt, saupbu_sales.collect_tot_amt, saupbu_sales.emp_name from saupbu_sales INNER JOIN sales_collect ON saupbu_sales.approve_no = sales_collect.approve_no where (collect_amt <> 0) and (collect_date >='"&from_date&"' and collect_date <= '"&to_date&"') "

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " ORDER BY emp_name, company, sales_date,collect_date, slip_no, collect_seq ASC"

Sql = "SELECT count(*) FROM saupbu_sales INNER JOIN sales_collect ON saupbu_sales.approve_no = sales_collect.approve_no where (collect_amt <> 0) and (collect_date >='"&from_date&"' and collect_date <= '"&to_date&"') " + field_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

for i = 0 to 4
	sum_tab(i,1) = 0
	sum_tab(i,2) = 0
next

sql = "select bill_collect, count(*), sum(collect_amt) as collect from saupbu_sales INNER JOIN sales_collect ON saupbu_sales.approve_no = sales_collect.approve_no where (collect_amt <> 0) and (collect_date >='"&from_date&"' and collect_date <= '"&to_date&"') " + field_sql + " group by bill_collect"
rs_sum.Open Sql, Dbconn, 1
do until rs_sum.eof
	if rs_sum(0) = "어음" then
		sum_tab(2,1)  = cdbl(rs_sum(1))
		sum_tab(2,2)  = cdbl(rs_sum(2))
	  elseif rs_sum(0) = "카드" then
		sum_tab(3,1)  = cdbl(rs_sum(1))
		sum_tab(3,2)  = cdbl(rs_sum(2))
	  elseif rs_sum(0) = "외환" then
		sum_tab(4,1)  = cdbl(rs_sum(1))
		sum_tab(4,2)  = cdbl(rs_sum(2))
	  else
		sum_tab(1,1)  = cdbl(rs_sum(1))
		sum_tab(1,2)  = cdbl(rs_sum(2))
	end if
	rs_sum.movenext()
loop
rs_sum.close()

for i = 1 to 4
	sum_tab(0,1) = sum_tab(0,1) + sum_tab(i,1)
	sum_tab(0,2) = sum_tab(0,2) + sum_tab(i,2)
next

Set rs_sum = Dbconn.Execute (sql)
if isnull(rs_sum("collect")) then
	tot_collect_amt = 0
  else
	tot_collect_amt = cdbl(rs_sum("collect"))
end if

sql = base_sql + field_sql + order_sql
Rs.Open Sql, Dbconn, 1

title_line = "수금 현황"
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
								<th scope="col">수금일자</th>
								<th scope="col">전표번호</th>
								<th scope="col">매출일자</th>
								<th scope="col">거래처명</th>
								<th scope="col">영업담당</th>
								<th scope="col">매출총액</th>
								<th scope="col">방법</th>
								<th scope="col">수금액</th>
								<th scope="col">수금총액</th>
								<th scope="col">잔액</th>
								<th scope="col">등록자</th>
								<th scope="col">등록일자</th>
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
								<td><%=rs("collect_date")%></td>
								<td><%=mid(rs("slip_no"),1,17)%></td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td><%=rs("bill_collect")%></td>
								<td class="right"><%=formatnumber(rs("collect_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("collect_tot_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_amt")-rs("collect_tot_amt"),0)%></td>
								<td><%=rs("reg_name")%></td>
								<td><%=rs("reg_date")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><strong><%=formatnumber(total_record,0)%>건<strong></td>
								<td colspan="12">
								<strong>현금</strong>&nbsp;&nbsp;<%=formatnumber(sum_tab(1,1),0)%>건&nbsp;,&nbsp;<%=formatnumber(sum_tab(1,2),0)%>원&nbsp;&nbsp;&nbsp;&nbsp;
								<strong>어음</strong>&nbsp;&nbsp;<%=formatnumber(sum_tab(2,1),0)%>건&nbsp;,&nbsp;<%=formatnumber(sum_tab(2,2),0)%>원&nbsp;&nbsp;&nbsp;&nbsp;
								<strong>카드</strong>&nbsp;&nbsp;<%=formatnumber(sum_tab(3,1),0)%>건&nbsp;,&nbsp;<%=formatnumber(sum_tab(3,2),0)%>원&nbsp;&nbsp;&nbsp;&nbsp;
								<strong>외환</strong>&nbsp;&nbsp;<%=formatnumber(sum_tab(4,1),0)%>건&nbsp;,&nbsp;<%=formatnumber(sum_tab(4,2),0)%>원
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

