<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

trade_code = request("trade_code")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT * FROM trade_person where trade_code = '"&trade_code&"' ORDER BY person_name ASC" 
Rs.Open Sql, Dbconn, 1

title_line = "거래처 담당자 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				return true;
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
							<col width="15%" >
							<col width="10%" >
							<col width="20%" >
							<col width="20%" >
							<col width="*" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">담당자</th>
								<th scope="col">직위</th>
								<th scope="col">전화번호</th>
								<th scope="col">계산서메일</th>
								<th scope="col">메모</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
	           			%>
							<tr>
								<td class="first"><%=rs("person_name")%></td>
								<td><%=rs("person_grade")%>&nbsp;</td>
								<td><%=rs("person_tel_no")%>&nbsp;</td>
								<td><%=rs("person_email")%>&nbsp;</td>
								<td><%=rs("person_memo")%>&nbsp;</td>
								<td><a href="#" onClick="pop_Window('trade_person_add.asp?trade_code=<%=rs("trade_code")%>&u_type=<%="U"%>','trade_person_add_pop','scrollbars=yes,width=800,height=220')">변경</a></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('trade_person_add.asp?trade_code=<%=trade_code%>','trade_person_add_pop','scrollbars=yes,width=800,height=220')" class="btnType04">거래처담당자등록</a>
					</div>                  
                    </td>
			      </tr>
				</table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

