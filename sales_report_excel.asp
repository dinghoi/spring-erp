<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

sales_month = request("sales_month")
sales_saupbu = request("sales_saupbu")
field_check = request("field_check")
field_view = request("field_view")

sales_yymm = mid(sales_month,1,4) + "-" + mid(sales_month,5,2)

savefilename = sales_month + "월 매출 내역.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from saupbu_sales where (substring(sales_date,1,7) = '"&sales_yymm&"')"

if field_check = "total" then
	field_sql = " "
  else
	field_sql = " and ("&field_check&" like '%"&field_view&"%') "
end if
if sales_saupbu = "전체" then
	saupbu_sql = " "
  else
	saupbu_sql = " and (saupbu = '"&sales_saupbu&"') "
end if
	
order_sql = " ORDER BY sales_date ASC"

sql = base_sql + field_sql + saupbu_sql + order_sql
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>회계 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">매출일자</th>
								<th scope="col">매출회사</th>
								<th scope="col">영업사업부</th>
								<th scope="col">고객사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">그룹</th>
								<th scope="col">합계금액</th>
								<th scope="col">공급가액</th>
								<th scope="col">세액</th>
								<th scope="col">담당자</th>
								<th scope="col">품목명</th>
							</tr>
						</thead>
						<tbody>
						<%
						end_sw = "N"
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=rs("sales_date")%></td>
								<td><%=rs("sales_company")%></td>
								<td><%=rs("saupbu")%></td>
								<td><%=rs("company")%></td>
								<td><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("vat_amt"),0)%></td>
								<td><%=rs("emp_name")%>&nbsp;</td>
								<td class="left"><%=rs("sales_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

