<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
account_group = request("account_group")
account_seq = request("account_seq")
account_name = request("account_name")
item_seq = request("item_seq")

account_item = ""
cost_yn = "N"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_acc = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "출금전표 사용적요 등록"
if u_type = "U" then

	Sql="select * from account_item where account_group ='" + account_group + "' and account_seq ='" + account_seq + "' and item_seq ='" + item_seq + "'"
	Set rs=DbConn.Execute(Sql)

	account_name = rs("account_name")
	account_item = rs("account_item")
	cost_yn = rs("cost_yn")
	rs.close()

	title_line = "출금전표 사용적요 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.account_item.value == "") {
					alert('적요를 입력하세요');
					frm.account_item.focus();
					return false;}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="account_code_add_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">계정과목</th>
								<td class="left"><%=account_name%></td>
							</tr>
							<tr>
								<th>적 요</th>
								<td class="left"><input name="account_item" type="text" value="<%=account_item%>" style="width:200px" onKeyUp="checklength(this,50);"></td>
							</tr>
							<tr>
								<th>비용</th>
								<td class="left">
                                <input type="radio" name="cost_yn" value="Y" <% if cost_yn = "Y" then %>checked<% end if %> title="기본사용" style="width:40px" ID="Radio1">
								  기본사용
                                <input type="radio" name="cost_yn" value="C" <% if cost_yn = "C" then %>checked<% end if %> title="확장사용" style="width:40px" ID="Radio1">
								  확장사용
								<input type="radio" name="cost_yn" value="N" <% if cost_yn = "N" then %>checked<% end if %> title="미사용" style="width:40px" ID="Radio2">
								    미사용</td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
	                <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="account_group" value="<%=account_group%>" ID="Hidden1">
				<input type="hidden" name="account_seq" value="<%=account_seq%>" ID="Hidden1">
				<input type="hidden" name="account_name" value="<%=account_name%>" ID="Hidden1">
				<input type="hidden" name="item_seq" value="<%=item_seq%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

