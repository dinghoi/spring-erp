<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim rsMem, car_yn, rsCar, car_no, car_name, oil_kind, pass, hp
Dim title_line

objBuilder.Append "SELECT car_no, car_name, oil_kind FROM car_info WHERE owner_emp_no = '"&user_id&"' "

Set rsCar = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsCar.eof Or rsCar.bof Then
	car_no = ""
	car_name = ""
	oil_kind = ""
Else
	car_no = rsCar("car_no")
	car_name = rsCar("car_name")
	oil_kind = rsCar("oil_kind")
End If
rsCar.Close() : Set rsCar = Nothing

objBuilder.Append "SELECT car_yn, user_name, user_id, pass, hp FROM memb WHERE user_id = '"&user_id&"' "

Set rsMem = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsMem.EOF Or rsMem.BOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('���� ������ �� �� �����ϴ�.');"
	Response.Write "	parent.opener.location.reload();"
	Response.Write "	self.close();"
	Response.Write "</script>"
End If

user_name = rsMem("user_name")
user_id = rsMem("user_id")
pass = rsMem("pass")
hp = rsMem("hp")

If f_toString(rsMem("car_yn"), "") = "" Or rsMem("car_yn") = "N" Then
	car_yn = "N"
Else
	car_yn = "Y"
End If

rsMem.Close() : Set rsMem = Nothing
DBConn.Close() : Set DBConn = Nothing

title_line = "����� ���� ����"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>���� ���� ����</title>
	<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<script type="text/javascript">
		function goAction(){
		   window.close();
		}
		function frmcheck(){
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			/*k = 0;

			for (j=0;j<2;j++){
				if(eval("document.frm.car_yn[" + j + "].checked")){
					k = j + 1
				}
			}*/

			if(document.frm.pass.value != document.frm.re_pass.value){
				alert('���� ��й�ȣ�� ��ġ���� �ʽ��ϴ�.');
				frm.re_pass.focus();
				return false;
			}

			if(document.frm.mod_pass.value =="") {
				alert('�����й�ȣ�� �Է��ϼ���.');
				frm.mod_pass.focus();
				return false;
			}

			if(document.frm.mod_pass.value != document.frm.mod_re_pass.value){
				alert('����Ȯ�κ�й�ȣ�� ��ġ�����ʽ��ϴ�.');
				frm.mod_pass.focus();
				return false;
			}

			if(document.frm.hp.value ==""){
				alert('�ڵ��� ��ȣ�� �Է��ϼ���.');
				frm.hp.focus();
				return false;
			}

			/*if(document.frm.old_car_yn.value =="Y"){
				if(k==1){
					alert('���� ������ �������� �ʽ��ϱ�??');
				}
			}

			if(k==2){
				if(document.frm.car_no.value ==""){
					frm.car_no.focus();
					alert('������ȣ�� �Է��ϼ���');
					return false;
				}

				if(document.frm.car_name.value ==""){
					frm.car_name.focus();
					alert('������ �Է��ϼ���');
					return false;
				}

				if(document.frm.oil_kind.value =="") {
					frm.oil_kind.focus();
					alert('������ �Է��ϼ���');
					return false;
				}
			}*/

			var result = confirm('���� �Ͻðڽ��ϱ�?');

			if(result == true){
				return true;
			}
			return false;
		}

		function car_yn_view(){
			k = 0;

			for (j=0;j<2;j++) {
				if (eval("document.frm.car_yn[" + j + "].checked")) {
					k = j + 1
				}
			}

			if (k==1) {
				document.getElementById('car_no_view').style.display = 'none';
				document.getElementById('car_name_view').style.display = 'none';
				document.getElementById('oil_kind_view').style.display = 'none';
			}

			if (k==2) {
				document.getElementById('car_no_view').style.display = '';
				document.getElementById('car_name_view').style.display = '';
				document.getElementById('oil_kind_view').style.display = '';
			}
		}
	</script>

</head>
<!--<body onload="car_yn_view();">-->
<body>
	<div id="container">
		<h3 class="tit"><%=title_line%></h3><br/>
		<form action="/member/user_mod_ok.asp" method="post" name="frm">
		<div class="gView">
			<table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="30%" >
					<col width="*" >
				</colgroup>
				<tbody>
					<tr>
						<th class="first">�����</th>
						<td class="left"><%=user_name%>(<%=user_id%>)</td>
					</tr>
					<tr>
						<th class="first">������й�ȣ</th>
						<td class="left">
							<input type="password" name="re_pass" id="re_pass" style="width:150px"/>
							<input type="hidden" name="pass" id="pass" value="<%=pass%>"/>
						</td>
					</tr>
					<tr>
						<th class="first">�����й�ȣ</th>
						<td class="left">
							<input type="password" name="mod_pass" id="mod_pass" onKeyUp="checklength(this,15);" style="width:150px;"/>
						</td>
					</tr>
					<tr>
						<th class="first">����Ȯ�κ�й�ȣ</th>
						<td class="left">
							<input type="password" name="mod_re_pass" id="mod_re_pass" style="width:150px;"/>
						</td>
					</tr>
					<tr>
						<th class="first">�ڵ�����ȣ</th>
						<td class="left">
							<input type="text" name="hp" id="hp" value="<%=hp%>" style="width:150px;"/>
						</td>
					</tr>
					<tr>
						<th class="first">��������</th>
						<td class="left">
							<!--<input type="radio" name="car_yn" value="N" <%'If car_yn = "N" Then %>checked<%'End If %> style="width:25px"  onClick="car_yn_view();">�̺���
							<input type="radio" name="car_yn" value="Y" <%'If car_yn = "Y" Then %>checked<%'End If %> style="width:25px" onClick="car_yn_view();">����-->
						<%
						If car_yn = "Y" Then
							Response.Write "����"
						Else
							Response.Write "�̺���"
						End If
						%>
						</td>
					</tr>
					<tr id="car_no_view">
					  <th class="first">������ȣ</th>
					  <td class="left">
						<!--<input name="car_no" type="text" id="car_no" value="<%'=car_no%>" style="width:150px">-->
						<%=car_no%>
					  </td>
					</tr>
					<tr id="car_name_view">
					  <th class="first">����</th>
					  <td class="left">
						<!--<input name="car_name" type="text" id="car_name" value="<%=car_name%>" style="width:150px">-->
						<%=car_name%>
					  </td>
					</tr>
					<tr id="oil_kind_view">
					  <th class="first">����</th>
					  <td class="left">
						<!--<select name="oil_kind" id="oil_kind" style="width:150px">
						  <option value="">����</option>
						  <option value="�ֹ���" <%If oil_kind = "�ֹ���" then %>selected<% end if %>>�ֹ���</option>
						  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
						  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
						</select>-->
						<%=oil_kind%>
					  </td>
					</tr>
				</tbody>
			</table>
		</div>
		<br>
		<div align="center">
			<span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" /></span>
			<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();" /></span>
		</div>
			<input name="old_car_yn" type="hidden" id="old_car_yn" value="<%=car_yn%>" />
			<input name="old_car_no" type="hidden" id="old_car_no" value="<%=car_no%>" />
		</form>
	</div>
</body>
</html>

