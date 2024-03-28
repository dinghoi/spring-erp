<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
card_upjong = request("card_upjong")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from card_slip where upjong = '"&card_upjong&"' order by slip_date desc"
Rs.Open Sql, Dbconn, 1

title_line = "카드 사용 현황 " + "- 카드거래처 업종 : " + card_upjong 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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

        </script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="13%" >
							<col width="15%" >
							<col width="*" >
							<col width="11%" >
							<col width="11%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사용일</th>
								<th scope="col">카드번호</th>
								<th scope="col">부서명/사용자</th>
								<th scope="col">사용처</th>
								<th scope="col">계정과목</th>
								<th scope="col">항목</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
							</tr>
						</thead>
						<tbody>
					  	<%
                        do until rs.eof
                      	%>
							<tr>
								<td class="first"><%=rs("slip_date")%></td>
								<td><%=rs("card_no")%></td>
								<td><%=rs("org_name")%>/<%=rs("org_name")%></td>
								<td><%=rs("customer")%></td>
								<td><%=rs("account")%>&nbsp;</td>
								<td><%=rs("account_item")%>&nbsp;</td>
								<td><%=formatnumber(rs("price"),0)%></td>
								<td><%=formatnumber(rs("cost"),0)%></td>
								<td><%=formatnumber(rs("cost_vat"),0)%></td>
							</tr>
							<%
                                rs.movenext()
                            loop
                            %>
						</tbody>
					</table>                    
					<br>
				</form>
				</div>
			</div>
	</body>
</html>

