<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw

bill_month = request("bill_month")
owner_company = request("owner_company")
field_check = request("field_check")
field_view = request("field_view")

from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

savefilename = bill_month + "월 이세로 세금계산서 내역.xls"

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

base_sql = "select * from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (end_yn = 'Y') and (cost_reg_yn = 'N') and (bill_id ='1') "
	
if field_check = "total" then
	field_sql = " "
  else
	field_sql = " and ("&field_check&" like '%"&field_view&"%') "
end if
if owner_company = "전체" then
	owner_sql = " "
  else
	owner_sql = " and (owner_company = '"&owner_company&"') "
end if
	
order_sql = " ORDER BY bill_date ASC"

sql = base_sql + field_sql + owner_sql + order_sql
Rs.Open Sql, Dbconn, 1

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
							<col width="6%" >
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="6%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">발행일</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호명</th>
								<th scope="col">대표자명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">청구</th>
								<th scope="col">담당자</th>
								<th scope="col">공급받는자이메일</th>
								<th scope="col">거래내역</th>
							</tr>
						</thead>
						<tbody>
						<%
						end_sw = "N"
						do until rs.eof
							Sql="select * from trade where trade_no = '"&rs("trade_no")&"'"
							Set rs_trade=DbConn.Execute(Sql)
							trade_sw = "Y"
							if rs_trade.eof or rs_trade.bof then
								trade_sw = "N"
							end if
							if rs("receive_email") = "" or isnull(rs("receive_email")) then
								emp_name = "-"
								emp_saupbu = "-"
							  else							
								k = instr(1,rs("receive_email"),"@")
								if k < 2 or isnull(k) then
									k = 2
								end if
								Sql="select * from emp_master where emp_email = '"&mid(trim(rs("receive_email")),1,k-1)&"'"
								Set rs_emp=DbConn.Execute(Sql)
								if rs_emp.eof then
									emp_name = "-"
									emp_saupbu = "-"
								  else
									emp_name = rs_emp("emp_name")
									emp_saupbu = rs_emp("emp_saupbu")
								end if
							end if
						%>
							<tr>
								<td class="first"><%=rs("bill_date")%></td>
								<td><%=rs("owner_company")%></td>
								<td><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("trade_owner")%></td>
								<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("bill_collect")%></td>
								<td><%=emp_name%></td>
								<td><%=rs("receive_email")%></td>
								<td class="left"><%=rs("tax_bill_memo")%></td>
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

