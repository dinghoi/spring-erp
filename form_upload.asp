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

title_line = "회사별 양식 UPLOAD"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
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
					alert('양식명을 입력하세요 !!!');
					frm.form_name.focus();
					return false;}
				if(document.frm.up_file.value =="") {
					alert('업로드 파일을 선택하세요 !!!');
					frm.up_file.focus();
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
								<th class="first">회사명</th>
								<td class="left"><%=company%></td>
								<th>양식명</th>
								<td class="left"><input name="form_name" type="text" id="form_name" style="width:150px" onKeyUp="checklength(this,30)"></td>
							</tr>
							<tr>
								<th class="first">업로드파일</th>
								<td colspan="3" class="left"><input name="up_file" type="file" id="up_file" size="70"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="company" value="<%=company%>">
				<input type="hidden" name="seq" value="<%=seq%>">
			</form>
		</div>				
	</body>
</html>

