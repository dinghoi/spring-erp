<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
asset_no = request("asset_no")
asset_name = request("asset_name")
asset_company = mid(asset_no,1,2)
install_date = mid(now(),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect
sql = "select * from asset where asset_no = '"&asset_no&"'"
set rs=dbconn.execute(sql)

title_line = "�ڻ� ��ġ ���"
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=install_date%>" );
			});	  
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

				if(document.frm.install_date.value == "") {
					alert('ó�����ڸ� �Է��ϼ���!!');
					frm.install_date.focus();
					return false;}
				{
				a=confirm('����� �Ͻðڽ��ϱ�?')
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
				<form action="asset_inst_reg_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">�ڻ��ȣ</th>
								<td class="left"><%=mid(asset_no,1,2)%>-<%=mid(asset_no,3,6)%>-<%=right(asset_no,4)%><input name="asset_no" type="hidden" id="asset_no" value="<%=asset_no%>"></td>
							</tr>
							<tr>
								<th class="first">�ڻ��</th>
								<td class="left"><%=asset_name%></td>
							</tr>
							<tr>
								<th class="first">��ġ����</th>
								<td class="left">
								<input name="inst_process" type="radio" value="Y" checked>��ġ 
              					<input type="radio" name="inst_process" value="R">�ݼ�
                                </td>
							</tr>
							<tr>
								<th class="first">�����</th>
								<td class="left"><input name="user_name" type="text" id="user_name" value="<%=rs("user_name")%>" style="width:150px" onKeyUp="checklength(this,20)"></td>
							</tr>
							<tr>
								<th class="first">ó������</th>
								<td class="left">
                                  <input name="install_date" type="text" value="<%=install_date%>" style="width:70px" id="datepicker">
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

