<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

emp_no=Request("emp_no")
from_date=Request("from_date")
to_date=Request("to_date")

title_line = user_name + "님 카드 전표 내역"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from card_slip where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and emp_no ='"&emp_no&"' ORDER BY slip_date ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<style type="text/css">
    <!--
    	.style10 {font-size: 10px; font-family: "굴림체", "굴림체", Seoul; }
        .style10B {font-size: 10px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; }
    -->
    </style>
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="6%" >
							<col width="12%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
							<col width="2%" >
							<col width="7%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr class="style10B">
								<th class="first" scope="col">카드사명</th>
								<th scope="col">카드번호</th>
								<th scope="col">승인일자</th>
								<th scope="col">사업자번호</th>
								<th scope="col">거래처명</th>
								<th scope="col">거래처유형</th>
								<th scope="col">공급가액</th>
								<th scope="col">세액</th>
								<th scope="col">합계금액</th>
								<th scope="col">부가세공제여부</th>
								<th scope="col">부가세유형</th>
								<th scope="col">계정과목</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						j = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1
							if rs("cost_vat") > 0 then
								vat_yn = "공제"
							  else
								vat_yn = "불공제"
							end if

							Sql="select * from account where account_name = '" + rs("account") + "'"
							Set rs_acc=DbConn.Execute(Sql)
						%>
							<tr class="style10">
								<td class="first"><%=rs("card_type")%></td>
								<td><%=rs("card_no")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("customer_no")%></td>
								<td><%=rs("customer")%></td>
								<td><%=rs("upjong")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td><%=vat_yn%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("account_item")%></td>
							</tr>
					  <%
							rs.movenext()
						loop
						rs.close()
						if price_sum <> ( cost_sum + cost_vat_sum ) then
							err_msg = "금액확인 요망"
						  else
						  	err_msg = " "
						end if
						%>
							<tr class="style10B">
								<th colspan="1" class="first">총계</th>
								<th colspan="4"><%=i%>&nbsp;건</th>
								<td></td>
							  	<th><%=formatnumber(cost_sum,0)%></th>
								<th><%=formatnumber(cost_vat_sum,0)%></th>
							  	<th><%=formatnumber(price_sum,0)%></th>
								<th colspan="3">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

