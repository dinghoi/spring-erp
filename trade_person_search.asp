<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
trade_code = Request("trade_code")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

SQL = "select * from trade_person where trade_code = '" + trade_code + "' ORDER BY person_name ASC"
Rs.open SQL, Dbconn, 1

title_line = "거래처 담당자 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>거래처 담당자 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function trade_person_list(person_name,person_email,person_tel_no)
			{
				opener.document.frm.trade_person.value = person_name;
				opener.document.frm.trade_person_tel_no.value = person_tel_no;
				opener.document.frm.trade_email.value = person_email;
				window.close();
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
				<form action="trade_person_search.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="25%" >
							<col width="20%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">담당자</th>
								<th scope="col">이메일</th>
								<th scope="col">연락처</th>
								<th scope="col">메모</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="trade_person_list('<%=rs("person_name")%>','<%=rs("person_email")%>','<%=rs("person_tel_no")%>');"><%=rs("person_name")%></a>
                                </td>
								<td><%=rs("person_email")%>&nbsp;</td>
								<td><%=rs("person_tel_no")%>&nbsp;</td>
								<td><%=rs("person_memo")%>&nbsp;</td>
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

