<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
acpt_no = request("acpt_no")

Sql="select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs=DbConn.Execute(Sql)
old_date = mid(rs("acpt_date"),1,10)

title_line = "�������� ����"

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
												$( "#datepicker" ).datepicker("setDate", "<%=acpt_date%>" );
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
				if(document.frm.acpt_date.value == "") {
					alert('���� �������ڸ� �Է��ϼ���');
					frm.acpt_date.focus();
					return false;}
				if(document.frm.acpt_date.value => document.frm.old_date.value) {
					alert('�߻����ڰ� �����Ϻ��� Ŭ���� �����ϴ�.');
					frm.acpt_date.focus();
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
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="acpt_date_mod_ok.asp" method="post" name="frm">
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
				        <th class="first">��������</th>
				        <td class="left"><%=rs("acpt_date")%>
				        <input name="old_date" type="hidden" value="<%=old_date%>">
                        </td>
				        <th>�����</th>
				        <td class="left"><%=rs("acpt_user")%></td>
			          </tr>
				      <tr>
				        <th class="first">ȸ��</th>
				        <td class="left"><%=rs("company")%></td>
				        <th>������</th>
				        <td class="left"><%=rs("dept")%></td>
			          </tr>
				      <tr>
				        <th class="first">ó������</th>
				        <td class="left"><%=rs("as_type")%></td>
				        <th>ó����Ȳ</th>
				        <td class="left"><%=rs("as_process")%></td>
			          </tr>
				      <tr>
				        <th class="first">���CE</th>
				        <td class="left"><%=rs("mg_ce")%>(<%=rs("mg_ce_id")%>)</td>
				        <th>��������</th>
				        <td class="left"><input name="acpt_date" type="text" id="datepicker" style="width:80px;text-align:center" value="<%=old_date%>" readonly="true"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                  <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
                    <input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

