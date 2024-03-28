<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

min_month = "201501"
now_month = cstr(mid(now(),1,4)) + cstr(mid(now(),6,2))

title_line = "비용 마감 일괄 취소"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			function getPageCode(){
				return "1 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
				if(document.frm.from_month.value =="") {
					alert('from년월을 입력하세요');
					frm.from_month.focus();
					return false;}
				if(document.frm.to_month.value =="") {
					alert('to년월을 입력하세요');
					frm.to_month.focus();
					return false;}
				if(document.frm.from_month.value < document.frm.min_month.value) {
					alert('from 년월은 201501보다 같거나 커야 한다');
					frm.from_month.focus();
					return false;}
				if(document.frm.to_month.value >= document.frm.now_month.value) {
					alert('to 년월은 현재년월 보다 작아야 한다 한다');
					frm.to_month.focus();
					return false;}

				a=confirm('처리 조건이 맞고? 정말 취소하시겠습니까???');
				if (a==true) {
					return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_end_condi_cancel_ok.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>처리조건</dt>
                        <dd>
                            <p>
                                <label>
                                    &nbsp;&nbsp;<strong>FROM년월&nbsp;</strong> :
                                    <input name="from_month" type="text" value="<%=from_month%>" style="width:70px" maxlength="6">
                                    &nbsp;~&nbsp;
                                    &nbsp;&nbsp;<strong>TO년월&nbsp;</strong> :
                                    <input name="to_month" type="text" value="<%=to_month%>" style="width:70px" maxlength="6">
                                </label>
                                    &nbsp;&nbsp;년월의 예)201501

                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                    <input type="hidden" name="min_month" value="<%=min_month%>" ID="Hidden1">
                    <input type="hidden" name="now_month" value="<%=now_month%>" ID="Hidden1">
				</form>
		</div>
	</div>
	</body>
</html>

