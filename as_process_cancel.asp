<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
acpt_no = request("acpt_no")

Sql="select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs=DbConn.Execute(Sql)

if rs("work_man_cnt") = "" or isnull(rs("work_man_cnt")) or rs("work_man_cnt") = 0 then
	etc_man_cnt = 0
  else
 	etc_man_cnt = int(rs("work_man_cnt")) - 1 	
end if  

if rs("overtime") = "Y" then
	overtime_view = "��Ư���Է�"
  else
  	overtime_view = "�Է¾���"
end if

Sql = "SELECT count(*) FROM att_file where acpt_no = "&int(acpt_no)
Set RsCount = Dbconn.Execute (sql)

att_cnt = cint(RsCount(0)) 'Result.RecordCount

title_line = "�Ϸ� ���"

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

				if(document.frm.cancel_pass.value != "123456") {
					alert('��� ��й�ȣ�� �ٸ��ϴ�.');
					frm.cancel_pass.focus();
					return false;}

				{
				a=confirm('����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_process_cancel_ok.asp" method="post" name="frm">
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
				        <td class="left"><%=rs("acpt_date")%><input name="old_date" type="hidden" id="old_date" value="<%=old_date%>"></td>
				        <th>�����</th>
				        <td class="left"><%=rs("acpt_user")%></td>
			          </tr>
				      <tr>
				        <th class="first">ȸ��/������</th>
				        <td class="left"><%=rs("company")%>&nbsp;<%=rs("dept")%></td>
				        <th>ó������</th>
				        <td class="left"><%=rs("as_type")%></td>
			          </tr>
				      <tr>
				        <th class="first">�۾��η�</th>
				        <td class="left"><%=rs("mg_ce")%>(<%=rs("mg_ce_id")%>)�� <%=etc_man_cnt%>&nbsp;��</td>
				        <th>÷������</th>
				        <td class="left"><%=att_cnt%>&nbsp;��</td>
			          </tr>
				      <tr>
				        <th class="first">��Ư���Է�</th>
				        <td class="left"><%=overtime_view%></td>
				        <th>��Һ��</th>
				        <td class="left"><input name="cancel_pass" type="password" id="cancel_pass" style="width:100px;text-align:center"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                <span class="btnType01"><input type="button" value="�Ϸ����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="Hidden1">
                <input type="hidden" name="overtime" value="<%=rs("overtime")%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

