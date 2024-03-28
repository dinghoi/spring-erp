<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

cost_month = request("cost_month")
emp_no = request("emp_no")

sql = "select * from person_cost where cost_month = '"&cost_month&"' and emp_no = '"&emp_no&"'"
Set rs=DbConn.Execute(Sql)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
				if(document.frm.variation_memo.value =="") {
					alert('증감사유를 입력하세요');
					frm.variation_memo.focus();
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
				<h3 class="tit">증감사유 등록 및 변경</h3>
				<form action="variation_memo_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="18%" >
				      <col width="32%" >
				      <col width="18%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">발생년월</th>
				        <td class="left"><%=mid(cost_month,1,4)%>년<%=mid(cost_month,5)%>월</td>
				        <th>담당자</th>
				        <td class="left"><%=rs("emp_name")%>(<%=rs("emp_no")%>)</td>
			          </tr>
				      <tr>
				        <th class="first">증감사유</th>
				        <td colspan="3" class="left">
                        <textarea name="variation_memo" id="textarea" rows="10"><%=rs("variation_memo")%></textarea>
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                </div>
				<br>
                    <input type="hidden" name="cost_month" value="<%=cost_month%>" ID="Hidden1">
                    <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

