<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
slip_id = request("slip_id")
slip_no = request("slip_no")
slip_seq = request("slip_seq")

Sql="select * from sales_slip where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
Set rs=DbConn.Execute(Sql)

title_line = "계산서 발행일 등록"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=bill_issue_date%>" );
			});	  
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
				if(document.frm.bill_issue_date.value == "") {
					alert('계산서 발행일을 입력하세요');
					frm.bill_issue_date.focus();
					return false;}
				if(document.frm.sales_date.value > document.frm.bill_issue_date.value) {
					alert('계산서 발행일이 매출일자보다 작을 수 없습니다.');
					frm.bill_issue_date.focus();
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
				<form action="sales_bill_issue_date_ok.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="30%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">계산서 발행일</th>
				        <td class="left"><input name="bill_issue_date" type="text" value="<%=bill_issue_date%>" style="width:80px;text-align:center" id="datepicker"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
					<input type="hidden" name="sales_date" value="<%=rs("sales_date")%>">
					<input type="hidden" name="slip_id" value="<%=slip_id%>">
					<input type="hidden" name="slip_no" value="<%=slip_no%>">
					<input type="hidden" name="slip_seq" value="<%=slip_seq%>">
			</form>
		</div>				
	</body>
</html>

