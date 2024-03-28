<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
gubun = Request("gubun")

trade_name = Request.form("trade_name")
if gubun = "" or isnull(gubun) then
	gubun = Request.form("gubun")
end if
'response.write(gubun)
Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

If user_id = "100359" Or user_id = "102592" Then
	condiSql = ""
Else
	condiSql = "AND trade_name NOT IN ('기타사업부', '기타') "
End If

SQL = "select * from trade where trade_name "
if trade_name = "" then
	'SQL = "select * from trade where trade_name = '" + trade_name + "' ORDER BY trade_name ASC"
	SQL = SQL&"= '"&trade_name&"' "&condiSQL
else
	'SQL = "select * from trade where trade_name like '%" + trade_name + "%' ORDER BY trade_name ASC"
	SQL = SQL&"LIKE '%"&trade_name&"%' "&condiSQL
end If
SQL = SQL&"ORDER BY trade_name ASC"

Rs.open SQL, Dbconn, 1

title_line = "거래처 검색"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>거래처 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  <script src="/java/jquery-1.9.1.js"></script>
	  <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function trade_list(trade_code,trade_name,trade_no,trade_person,trade_email)
			{
				var saupbu = '<%=saupbu %>';

				if (trade_name == '기타사업부') {
					//alert(saupbu);

					if (saupbu != '경영지원실' )	{
						alert("기타사업부는 경영지원실외 선택 할 수 없습니다.");
						return;
					}
				}

				if(document.frm.gubun.value =="1") {
					opener.document.frm.trade_code.value = trade_code;
					opener.document.frm.trade_name.value = trade_name;
					opener.document.frm.trade_no.value = trade_no;
					opener.document.frm.trade_person.value = trade_person;
					opener.document.frm.trade_email.value = trade_email;
					window.close();
				}
				if(document.frm.gubun.value =="2") {
					opener.document.frm.bill_trade_name.value = trade_name;
					opener.document.frm.bill_trade_code.value = trade_code;
					window.close();
				}
				if(document.frm.gubun.value =="3") {
					opener.document.frm.customer.value = trade_name;
					opener.document.frm.customer_no.value = trade_no;
					window.close();
				}
				if(document.frm.gubun.value =="4") {
					opener.document.frm.company.value = trade_name;
					window.close();
				}
				if(document.frm.gubun.value =="5") {
					opener.document.frm.group_name.value = trade_name;
					window.close();
				}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.trade_name.value =="") {
					alert('거래처명을 입력하세요');
					frm.trade_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="trade_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>조직명을 입력하세요 </strong>
								<label>
        						<input name="trade_name" type="text" id="trade_name" value="<%=trade_name%>" style="width:150px;text-align:left;ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처명</th>
								<th scope="col">사업자번호</th>
								<th scope="col">담당자</th>
								<th scope="col">이메일</th>
								<th scope="col">담당본부</th>
							</tr>
						</thead>
						<tbody>
						<% if gubun = "2" or gubun = "5" then	%>
							<tr>
								<td class="first">
                                <a href="#" onClick="trade_list('','','','','');">없음</a>
                                </td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<% end if	%>
						<%
						i = 0
						do until rs.eof or rs.bof
							trade_code = rs("trade_code")
							trade_name = rs("trade_name")
							trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6)
							trade_saupbu = rs("saupbu")

							Sql="select * from trade_person where trade_code = '"&trade_code&"'"
							Set rs_etc=DbConn.Execute(Sql)
							if rs_etc.eof or rs_etc.bof then
								trade_person = ""
								trade_email = ""
							  else
								trade_person = rs_etc("person_name")
								trade_email = rs_etc("person_email")
							end if
							rs_etc.close()
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="trade_list('<%=trade_code%>','<%=trade_name%>','<%=trade_no%>','<%=trade_person%>','<%=trade_email%>');"><%=rs("trade_name")%></a>
                                </td>
								<td><%=trade_no%>&nbsp;</td>
								<td><%=trade_person%>&nbsp;</td>
								<td><%=trade_email%>&nbsp;</td>
								<td><%=trade_saupbu%>&nbsp;</td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="4">내역이 없습니다</td>
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

