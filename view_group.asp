<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
company = Request("company")

sql = "select * from trade where group_name = '" + company + "' ORDER BY trade_name ASC"
Rs.open SQL, Dbconn, 1

title_line = "조회 그룹 내역"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>조회 그룹 내역</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
				<h3 class="stit">조회회사를 추가하려면 회사 코드 관리에서 회사를 선택하신 후 그룹에&nbsp;'<%=company%>'&nbsp;를 입력하시고 저장하시면 됩니다.</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="50%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">회사명</th>
								<th scope="col">그룹명</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
						%>
							<tr>
								<td class="first"><%=rs("trade_name")%></td>
								<td><%=rs("group_name")%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 or i = 1 then
						%>
							<tr>
								<td class="first" colspan="2">
                                조회그룹이 없거나 본인 회사 한개만 조회가 됩니다.
                                </td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
			</form>
		</div>        				
	</body>
</html>

