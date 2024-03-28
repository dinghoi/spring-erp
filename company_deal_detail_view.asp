<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
cost_year = Request("cost_year")
cost_mm = Request("cost_mm")
cost_month = cstr(cost_year) + right("0" + cstr(cost_mm),2)
from_date = cstr(cost_year) + "-" + cstr(cost_mm) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

sql = "select * FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and cost_center = '회사간거래' order by slip_date asc"
rs.Open sql, Dbconn, 1
'Response.write sql & "<br>"

title_line = "회사간 거래 세부 내역"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
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
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
							<col width="10%" >
							<col width="8%" >
							<col width="14%" >
							<col width="16%" >
							<col width="*" >
							<col width="9%" >
						</colgroup>
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
							<tr>
								<th colspan="7" class="first">합계</th>
								<th class="right"><%=formatnumber(cost_sum,0)%></th>
							</tr>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnCenter">
                    <a href="saupbu_sales_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%="회사간거래"%>" class="btnType04">매출액 엑셀다운로드</a>
                    <a href="company_deal_detail_excel.asp?cost_month=<%=cost_month%>" class="btnType04">비용 엑셀다운로드</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
				</div>				        				
	</form>
	</body>
</html>

