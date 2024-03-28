<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
trade_code = request("trade_code")

person_name = ""
person_grade = ""
person_tel_no = ""
person_email = ""
person_memo = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "거래처 담당자 등록"
if u_type = "U" then

	Sql="select * from trade_person where trade_code = '"&trade_code&"'"
	Set rs=DbConn.Execute(Sql)

	trade_code = rs("trade_code")
	person_name = rs("person_name")
	person_grade = rs("person_grade")
	person_tel_no = rs("person_tel_no")
	person_email = rs("person_email")
	person_memo = rs("person_memo")
	rs.close()

	title_line = "거래처 담당자 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//ENrs("customer_no")http://www.w3.org/TR/html4/loose.dtd">
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
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.person_name.value =="") {
					alert('담당자를 입력하세요');
					frm.person_name.focus();
					return false;}
				if(document.frm.person_email.value =="") {
					alert('계산서 메일을 입력하세요');
					frm.person_email.focus();
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
				<form action="trade_person_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">담당자</th>
				        <td class="left">
					<% if u_type = "U" then	%>
                        <%=person_name%><input name="person_name" type="hidden" id="person_name" style="width:200px;" value="<%=person_name%>">
					<%   else	%>
                        <input name="person_name" type="text" id="person_name" style="width:200px;" value="<%=person_name%>" onKeyUp="checklength(this,20);">
					<% end if	%>
                        </td>
				        <th>담당자 직급</th>
				        <td class="left"><input name="person_grade" type="text" id="person_grade" style="width:200px;" value="<%=person_grade%>" onKeyUp="checklength(this,20);"></td>
			          </tr>
				      <tr>
				        <th class="first">전화번호</th>
				        <td class="left"><input name="person_tel_no" type="text" id="person_tel_no" style="width:200px;" value="<%=person_tel_no%>" onKeyUp="checklength(this,20);"></td>
				        <th>계산서메일</th>
				        <td class="left"><input name="person_email" type="text" id="person_email" style="width:200px;" value="<%=person_email%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
				      <tr>
				        <th class="first">메모</th>
				        <td colspan="3" class="left"><input name="person_memo" type="text" id="person_memo" style="width:500px" value="<%=person_memo%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onClick="javascript:frmcheck();" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onClick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="trade_code" value="<%=trade_code%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

