<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
user_id = request("user_id")

user_name = ""
user_grade = ""
team = ""
hp = ""
email = ""
reside = ""
reside_place = ""
sms = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "CE ���"
if u_type = "U" then

	Sql="select * from memb where user_id = '" + user_id + "'"
	Set rs=DbConn.Execute(Sql)

	user_id = rs("user_id")
	user_name = rs("user_name")
	user_grade = rs("user_grade")
	hp = rs("hp")
	team = rs("team")
	email = rs("email")
	grade = rs("grade")
	reside = rs("reside")
	reside_place = rs("reside_place")
	sms = rs("sms")
	rs.close()

	title_line = "CE ����"
end if
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
			function menu1() {
			var c = document.frm.reside.options[document.frm.reside.selectedIndex].value;
				if (c == '0') {
					document.getElementById('reside_place').style.display = 'none';
				}
				if (c == '1') {
					document.getElementById('reside_place').style.display = '';
				}
			}
			function moveNext(varControl,varNext){
				var ctrType="";
			
				if(varControl.value.length == varControl.maxLength){
					varNext.focus();
					ctrType = varNext.type.toUpperCase();
					if(ctrType != "RADIO" && ctrType != "SELECT-ONE")
						varNext.select();
				}
			}

			function chkfrm() {
				if(document.frm.user_id.value =="") {
					alert('���̵� �Է��ϼ���');
					frm.user_id.focus();
					return false;}
				if(document.frm.team.value =="") {
					alert('�μ����� �����ϼ���');
					frm.team.focus();
					return false;}
				if(document.frm.hp.value =="") {
					alert('�ڵ��� ��ȣ�� �Է��ϼ���');
					frm.hp.focus();
					return false;}
				if(document.frm.reside.value == "1") {
					if(document.frm.reside_place.value == "����") {
						alert('�����̸� ����ó�� ���簡 �ɼ� ����.');
						frm.reside.focus();
						return false;}}				
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.sms[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("���ڹ߼� ���θ� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	

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
	<body oncontextmenu="return false" ondragstart="return false" onload="menu1()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="ce_reg_ok.asp" method="post" name="frm">
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
								<th class="first">���̵�</th>
								<td class="left">
                                <%	if u_type = "U" then %>
                                <%=user_id%><input type="hidden" name="user_id" value="<%=user_id%>">
                                <%	  else	%>
                                <input name="user_id" type="text" id="user_id" style="width:120px" readonly="true"><a href="#" class="btnType03" onclick="javascript:pop_id_check()" >���Ȯ��</a>
                                <% 	end if	%>
                                </td>
								<th>����ڸ�</th>
								<td class="left"><input name="user_name" type="text" id="user_name" style="width:120px" notnull errname="����ڸ�" onKeyUp="checklength(this,20)" value="<%=user_name%>"></td>
							</tr>
							<tr>
								<th>����</th>
								<td class="left">
								<%
                                    Sql="select * from etc_code where used_sw = 'Y' and etc_type = '61' order by etc_code asc"
                                    Rs_etc.Open Sql, Dbconn, 1
                                %>
                                <select name="user_grade" id="user_grade" style="width:120px">
                                <% 
                                    do until rs_etc.eof 
                                %>
                                        <option value="<%=rs_etc("etc_name")%>" <% if rs_etc("etc_name") = user_grade then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                <%
                                        rs_etc.movenext()  
                                    loop 
                                    rs_etc.Close()
                                %>
                                </select>
                                </td>
								<th>�μ���</th>
								<td class="left">
								<%
                                    Sql="select * from etc_code where used_sw = 'Y' and etc_type = '62' order by etc_code asc"
                                    Rs_etc.Open Sql, Dbconn, 1
                                %>
                                <select name="team" id="team" style="width:120px">
                                  	<option>����</option>
                                <% 
                                    do until rs_etc.eof 
                                %>
                                        <option value="<%=rs_etc("etc_name")%>" <% if rs_etc("etc_name") = team then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                <%
                                        rs_etc.movenext()  
                                    loop 
                                    rs_etc.Close()
                                %>
                                </select>
                                </td>
							</tr>
							<tr>
								<th class="first">�ڵ���</th>
								<td class="left"><input name="hp" type="text" id="hp" style="width:120px" onKeyUp="checklength(this,13);" value="<%=hp%>"></td>
								<th>�̸���</th>
								<td class="left"><input name="email" type="text" id="email" style="width:200px" onKeyUp="checklength(this,20)" value="<%=email%>"></td>
							</tr>
							<tr>
								<th class="first">�������</th>
								<td class="left">
                                <select name="grade" id="grade" style="width:120px">
                                  <option value="6" <% if grade = "6" then %>selected<% end if %>>���Ѵ��</option>
                                  <option value="5" <% if grade = "5" then %>selected<% end if %>>�����</option>
                                  <option value="4" <% if grade = "4" then %>selected<% end if %>>CE</option>
                                  <option value="3" <% if grade = "3" then %>selected<% end if %>>����CE</option>
                                  <option value="2" <% if grade = "2" then %>selected<% end if %>>���ְ�����</option>
                                  <option value="1" <% if grade = "1" then %>selected<% end if %>>������</option>
                                  <option value="0" <% if grade = "0" then %>selected<% end if %>>������</option>
                                </select>
                                </td>
								<th>�����׷�</th>
								<%
                                    Sql="select * from type_code where etc_type = '91' and etc_seq = '" + mg_group + "'"
                                    Rs_type.Open Sql, Dbconn, 1
                                %>
								<td class="left"><%=rs_type("type_name")%></td>
							</tr>
							<tr>
								<th class="first">���ֱ���</th>
								<td class="left">
                                  <select name="reside" id="reside" onChange="menu1()" style="width:70px">
                                    <option value="0" <% if reside = "0" then %>selected<% end if %>>�����</option>
                                    <option value="1" <% if reside = "1" then %>selected<% end if %>>����</option>
                                  </select>
                                  <%
                                        Sql="select * from etc_code where used_sw = 'Y' and mg_group = '" + mg_group + "' and etc_type = '55' order by etc_code asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="reside_place" id="reside_place" style="display:none; width:120px">
                                    <% 
                                        do until rs_etc.eof 
                                    %>
                                    <option value="<%=rs_etc("etc_name")%>" <% if rs_etc("etc_name") = reside_place then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                    <%
                                            rs_etc.movenext()  
                                        loop 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                </td>
								<th>���ڹ߼ۿ���</th>
								<td class="left">
									<input type="radio" name="sms" value="Y" <% if sms = "Y" then %>checked<% end if %> title="�߼�" style="width:40px" ID="Radio1">�߼�
									<input type="radio" name="sms" value="N" <% if sms = "N" then %>checked<% end if %> title="�߼۾���" style="width:40px" ID="Radio2">�߼۾���
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

