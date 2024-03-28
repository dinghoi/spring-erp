<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

min_month = "201501"
now_month = cstr(mid(now(),1,4)) + cstr(mid(now(),6,2))

title_line = "��� ���� �ϰ� ���"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
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
					alert('from����� �Է��ϼ���');
					frm.from_month.focus();
					return false;}
				if(document.frm.to_month.value =="") {
					alert('to����� �Է��ϼ���');
					frm.to_month.focus();
					return false;}
				if(document.frm.from_month.value < document.frm.min_month.value) {
					alert('from ����� 201501���� ���ų� Ŀ�� �Ѵ�');
					frm.from_month.focus();
					return false;}
				if(document.frm.to_month.value >= document.frm.now_month.value) {
					alert('to ����� ������ ���� �۾ƾ� �Ѵ� �Ѵ�');
					frm.to_month.focus();
					return false;}

				a=confirm('ó�� ������ �°�? ���� ����Ͻðڽ��ϱ�???');
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
					<legend>��ȸ����</legend>
					<dl>
						<dt>ó������</dt>
                        <dd>
                            <p>
                                <label>
                                    &nbsp;&nbsp;<strong>FROM���&nbsp;</strong> :
                                    <input name="from_month" type="text" value="<%=from_month%>" style="width:70px" maxlength="6">
                                    &nbsp;~&nbsp;
                                    &nbsp;&nbsp;<strong>TO���&nbsp;</strong> :
                                    <input name="to_month" type="text" value="<%=to_month%>" style="width:70px" maxlength="6">
                                </label>
                                    &nbsp;&nbsp;����� ��)201501

                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
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

