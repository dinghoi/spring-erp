<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
company = request("company")
seq = request("seq")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "ȸ�纰 ��� UPLOAD"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.form_name.value =="") {
					alert('��ĸ��� �Է��ϼ��� !!!');
					frm.form_name.focus();
					return false;}
				if(document.frm.up_file.value =="") {
					alert('���ε� ������ �����ϼ��� !!!');
					frm.up_file.focus();
					return false;}
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body onload="specview()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="form_upload_save.asp" method="post" name="frm" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
                          <col width="15%" >
                          <col width="35%" >
                          <col width="15%" >
                          <col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">ȸ���</th>
								<td class="left"><%=company%></td>
								<th>��ĸ�</th>
								<td class="left"><input name="form_name" type="text" id="form_name" style="width:150px" onKeyUp="checklength(this,30)"></td>
							</tr>
							<tr>
								<th class="first">���ε�����</th>
								<td colspan="3" class="left"><input name="up_file" type="file" id="up_file" size="70"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="company" value="<%=company%>">
				<input type="hidden" name="seq" value="<%=seq%>">
			</form>
		</div>				
	</body>
</html>

