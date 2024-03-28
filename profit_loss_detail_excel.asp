<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

cost_month = request("cost_month")

slip_month = mid(cost_month,1,4) + "-" + mid(cost_month,5,2)

title_line = cost_month + "월 세금계산서 내역"
savefilename = title_line + ".xls"

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

sql = "select * from general_cost where (tax_bill_yn = 'Y') and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
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
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">비용회사</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">조직명</th>
								<th scope="col">상주처</th>
								<th scope="col">담당자</th>
								<th scope="col">발행일자</th>
								<th scope="col">발행순번</th>
								<th scope="col">고객사</th>
								<th scope="col">담당영업사업부</th>
								<th scope="col">외주업체</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">비용유형</th>
								<th scope="col">비용구분</th>
								<th scope="col">세부유형</th>
								<th scope="col">발행내역</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("emp_company")%></td>
								<td><%=rs("bonbu")%></td>
								<td><%=rs("saupbu")%></td>
								<td><%=rs("team")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("reside_place")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("slip_seq")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("mg_saupbu")%></td>
								<td><%=rs("customer")%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("cost_center")%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
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

