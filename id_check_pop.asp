<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/army_dbcon.asp" -->
<%
user_id = Request.form("user_id")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

SQL = "select * from sla_user where user_id = '" + user_id + "'"
set rs=dbconn.execute(sql)

if user_id = "" or user_id = null then
	use_msg = "���̵� �Է��ϼ��� !!!!"
	use_ok = "N"
 else
	If Rs.Eof or Rs.Bof Then
		use_msg = "��밡���� ���̵��Դϴ�. ����Ͻðڽ��ϱ�?"
		use_ok = "Y"
	 else
		use_msg = "�̹� ����ϰ� �ִ� ���̵��Դϴ� !!!!"
		use_ok = "N"
	End if
end if

title_line = "���̵� �ߺ� Check"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>SLA ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function id_move(mg_ce_id)
			{
				opener.document.frm.user_id.value = user_id;
				window.close();
			
			}
			
			function chkfrm(){
				if(document.frm.user_id.value =="") {
					alert('���̵� �Է��ϼ���');
					frm.user_id.focus();
					return false;}
				document.frm.submit ();
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="id_check_pop.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="25%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">���̵�</th>
								<td class="left"><input name="user_id" type="text" id="user_id" value="<%=user_id%>" style="width:120px" onKeyUp="checklength(this,15)"><a href="#" class="btnType03" onclick="javascript:chkfrm();" >�ߺ�Ȯ��</a></td>
							</tr>
							<tr>
								<th class="first">���ɿ���</th>
								<td class="left"><%=use_msg%>&nbsp;
								  <% if use_ok = "Y" then %>
                                    <a href="#" class="btnType03" onClick="id_move('<%=user_id%>');">���</a>
                                  <% end if %>
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
		</div>				
	</body>
</html>

