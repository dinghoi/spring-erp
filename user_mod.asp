<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Set dbconn = server.CreateObject("adodb.connection")
dbconn.open dbconnect

Sql="select * from memb where user_id='"&user_id&"'"
Set Rs=dbconn.execute(Sql)
if rs.eof or rs.bof then
	response.write"<script language=javascript>"
	response.write"alert('���������� �� �� �����ϴ�');"		
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
end if
if rs("car_yn") = "" or isnull(rs("car_yn")) or rs("car_yn") = "N" then
	car_yn = "N" 
  else
	car_yn = "Y"
end if

sql = "select * from car_info where owner_emp_no = '"&user_id&"'"
Set rs_car=dbconn.execute(Sql)
if rs_car.eof or rs_car.bof then
	car_no = ""
	car_name = ""
	oil_kind = ""
  else
	car_no = rs_car("car_no")
	car_name = rs_car("car_name")
	oil_kind = rs_car("oil_kind")
end if

title_line = "����� ���� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			

			function chkfrm() {
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.car_yn[" + j + "].checked")) {
						k = j + 1
					}
				}
				if(document.frm.pass.value != document.frm.re_pass.value) {
					alert('��й�ȣ�� �ٸ��ϴ�.');
					frm.re_pass.focus();
					return false;}
//				if(document.frm.mod_pass.value =="") {
//					alert('�����й�ȣ�� �Է��ϼ���');
//					frm.mod_pass.focus();
//					return false;}
				if(document.frm.mod_pass.value != document.frm.mod_re_pass.value) {
					alert('���� Ȯ�� ��й�ȣ�� �ٸ��ϴ�');
					frm.mod_pass.focus();
					return false;}
				if(document.frm.hp.value =="") {
					alert('�ڵ��� ��ȣ�� �Է��ϼ���');
					frm.hp.focus();
					return false;}
				if(document.frm.old_car_yn.value =="Y") {
					if(k==1) {
						alert('���� ������ �������� �ʽ��ϱ�??');
						}}
				if(k==2) {
					if(document.frm.car_no.value =="") {
						frm.car_no.focus();
						alert('������ȣ�� �Է��ϼ���');
						return false;}}
				if(k==2) {
					if(document.frm.car_name.value =="") {
						frm.car_name.focus();
						alert('������ �Է��ϼ���');
						return false;}}
				if(k==2) {
					if(document.frm.oil_kind.value =="") {
						frm.oil_kind.focus();
						alert('������ �Է��ϼ���');
						return false;}}

				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function car_yn_view() 
			{
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.car_yn[" + j + "].checked")) {
						k = j + 1
					}
				}
				if (k==1) {
					document.getElementById('car_no_view').style.display = 'none'; 
					document.getElementById('car_name_view').style.display = 'none'; 
					document.getElementById('oil_kind_view').style.display = 'none'; }
				if (k==2) {
					document.getElementById('car_no_view').style.display = ''; 
					document.getElementById('car_name_view').style.display = ''; 
					document.getElementById('oil_kind_view').style.display = ''; }
			}
		</script>

	</head>
	<body onload="car_yn_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="user_mod_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">�����</th>
								<td class="left"><%=rs("user_name")%>(<%=rs("user_id")%>)</td>
							</tr>
							<tr>
								<th class="first">������й�ȣ</th>
								<td class="left"><input name="re_pass" type="password" id="re_pass" style="width:150px"><input name="pass" type="hidden" id="pass" value="<%=rs("pass")%>"></td>
							</tr>
							<tr>
								<th class="first">�����й�ȣ</th>
								<td class="left"><input name="mod_pass" type="password" id="mod_pass" onKeyUp="checklength(this,15);" style="width:150px"></td>
							</tr>
							<tr>
								<th class="first">����Ȯ�κ�й�ȣ</th>
								<td class="left"><input name="mod_re_pass" type="password" id="mod_re_pass" style="width:150px"></td>
							</tr>
							<tr>
								<th class="first">�ڵ�����ȣ</th>
								<td class="left"><input name="hp" type="text" id="hp" value="<%=rs("hp")%>" style="width:150px"></td>
							</tr>
							<tr>
								<th class="first">��������</th>
								<td class="left">
                                <input type="radio" name="car_yn" value="N" <% if car_yn = "N" then %>checked<% end if %> style="width:25px"  onClick="car_yn_view()">�̺���
								<input type="radio" name="car_yn" value="Y" <% if car_yn = "Y" then %>checked<% end if %> style="width:25px" onClick="car_yn_view()">����
                                </td>
            				</tr>
							<tr id="car_no_view">
							  <th class="first">������ȣ</th>
							  <td class="left"><input name="car_no" type="text" id="car_no" value="<%=car_no%>" style="width:150px"></td>
					        </tr>
							<tr id="car_name_view">
							  <th class="first">����</th>
							  <td class="left"><input name="car_name" type="text" id="car_name" value="<%=car_name%>" style="width:150px"></td>
					        </tr>
							<tr id="oil_kind_view">
							  <th class="first">����</th>
							  <td class="left">
                                <select name="oil_kind" id="oil_kind" style="width:150px">
								  <option value="">����</option>
								  <option value="�ֹ���" <%If oil_kind = "�ֹ���" then %>selected<% end if %>>�ֹ���</option>
								  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
								  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
							    </select>
                              </td>
					        </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
	                <input name="old_car_yn" type="hidden" id="old_car_yn" value="<%=car_yn%>">
	                <input name="old_car_no" type="hidden" id="old_car_no" value="<%=car_no%>">
				</form>
		</div>				
	</body>
</html>

