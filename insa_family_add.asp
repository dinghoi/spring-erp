<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
family_empno = request("family_empno")
family_seq = request("family_seq")
emp_name = request("emp_name")

family_rel = ""
family_name = ""
family_birthday = ""
family_birthday_id = ""
family_job = ""
family_live = "����"
family_person1 = ""
family_person2 = ""
family_tel_ddd = ""
family_tel_no1 = ""
family_tel_no2 = ""
family_support_yn = "N"
family_national = "������"
family_disab = ""
family_merit = ""
family_serius = ""
family_pensioner = ""
family_witak = ""
family_holt = ""
family_holt_date = ""
family_children = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " �������� ��� "
if u_type = "U" then

	Sql="select * from emp_family where family_empno = '"&family_empno&"' and family_seq = '"&family_seq&"'"
	Set rs=DbConn.Execute(Sql)

	family_rel = rs("family_rel")
    family_name = rs("family_name")
    family_birthday = rs("family_birthday")
    family_birthday_id = rs("family_birthday_id")
    family_job = rs("family_job")
    family_live = rs("family_live")
    family_person1 = rs("family_person1")
    family_person2 = rs("family_person2")
	family_tel_ddd = rs("family_tel_ddd")
    family_tel_no1 = rs("family_tel_no1")
    family_tel_no2 = rs("family_tel_no2")
	family_support_yn = rs("family_support_yn")
	if family_birthday = "1900-01-01"  then
	   family_birthday = ""
	end if
	family_national = rs("family_national")
    family_disab = rs("family_disab")
	family_merit = rs("family_merit")
    family_serius = rs("family_serius")
    family_pensioner = rs("family_pensioner")
    family_witak = rs("family_witak")
    family_holt = rs("family_holt")
    family_holt_date = rs("family_holt_date")
	if family_holt_date = "1900-01-01"  then
	   family_holt_date = ""
	end if
	family_children = rs("family_children")
	
	rs.close()

	title_line = " �������� ���� "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ�޿� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
			});	
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=family_holt_date%>" );
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
				if(document.frm.family_birthday.value =="") {
					alert('��������� �Է��ϼ���');
					frm.family_birthday.focus();
					return false;}
//				if(document.frm.family_person1.value =="") {
//					alert('�ֹε�Ϲ�ȣ�� �Է��ϼ���');
//					frm.family_person1.focus();
//					return false;}
//				if(document.frm.family_person2.value =="") {
//					alert('�ֹε�Ϲ�ȣ�� �Է��ϼ���');
//					frm.family_person2.focus();
//					return false;}
				if(document.frm.family_rel =="") {
					alert('�����׸��� �����ϼ���');
					frm.family_rel.focus();
					return false;}
				if(document.frm.family_name.value =="") {
					alert('���������� �Է��ϼ���');
					frm.family_name.focus();
					return false;}
				if(document.frm.family_tel_no1.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.family_tel_no1.focus();
					return false;}
				if(document.frm.family_tel_no2.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.family_tel_no2.focus();
					return false;}
				if(document.frm.family_support_yn.value =="") {
					alert('�ξ簡�����θ� �Է��ϼ���');
					frm.family_support_yn.focus();
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
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_family_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="15%" >
						<col width="18%" >
						<col width="15%" >
						<col width="18%" >
						<col width="15%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">���</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="family_empno" type="text" id="family_empno" size="14" value="<%=family_empno%>" readonly="true">
                      <input type="hidden" name="family_seq" value="<%=family_seq%>" ID="Hidden1"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>����(�ʼ�)</th>
                      <td colspan="5" class="left">
					  <select name="family_rel" id="family_rel" value="<%=family_rel%>" style="width:100px">
				          <option value="" <% if family_rel = "" then %>selected<% end if %>>����</option>
				          <option value='��' <%If family_rel = "��" then %>selected<% end if %>>��</option>
				          <option value='��' <%If family_rel = "��" then %>selected<% end if %>>��</option>
				          <option value='����' <%If family_rel = "����" then %>selected<% end if %>>����</option>
                          <option value='�Ƴ�' <%If family_rel = "�Ƴ�" then %>selected<% end if %>>�Ƴ�</option>
                          <option value='�Ƶ�' <%If family_rel = "�Ƶ�" then %>selected<% end if %>>�Ƶ�</option>
                          <option value='��' <%If family_rel = "��" then %>selected<% end if %>>��</option>
                          <option value='����' <%If family_rel = "����" then %>selected<% end if %>>����</option>
                          <option value='����' <%If family_rel = "����" then %>selected<% end if %>>����</option>
                          <option value='������' <%If family_rel = "������" then %>selected<% end if %>>������</option>
                          <option value='������' <%If family_rel = "������" then %>selected<% end if %>>������</option>
                          <option value='�ú�' <%If family_rel = "�ú�" then %>selected<% end if %>>�ú�</option>
                          <option value='�ø�' <%If family_rel = "�ø�" then %>selected<% end if %>>�ø�</option>
                          <option value='����' <%If family_rel = "����" then %>selected<% end if %>>����</option>
                          <option value='���' <%If family_rel = "���" then %>selected<% end if %>>���</option>
                          <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" then %>selected<% end if %>>��(�����ڸ�)</option>
                          <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" then %>selected<% end if %>>��(�����ڸ�)</option>
                          <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" then %>selected<% end if %>>��(�����ڸ�)</option>
                          <option value='����' <%If family_rel = "����" then %>selected<% end if %>>����</option>
                          <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" then %>selected<% end if %>>��(�����ڸ�)</option>
                          <option value='����' <%If family_rel = "����" then %>selected<% end if %>>����</option>
                          <option value='�ճ�' <%If family_rel = "�ճ�" then %>selected<% end if %>>�ճ�</option>
                          <option value='�ں�' <%If family_rel = "�ں�" then %>selected<% end if %>>�ں�</option>
                          <option value='�պ�' <%If family_rel = "�պ�" then %>selected<% end if %>>�պ�</option>
                          <option value='��Ÿ����' <%If family_rel = "��Ÿ����" then %>selected<% end if %>>��Ÿ����</option>
                      </select>
                      &nbsp;��Ź�Ƶ��ΰ��� ��Ÿ���踦 �����Ͻʽÿ�!
                      </td>
                    </tr>
                    <tr>
                      <th>����(�ʼ�)</th>
                      <td colspan="2" class="left">
					  <input name="family_name" type="text" id="family_name" size="14" value="<%=family_name%>"></td>
                      <th>�����</th>
                      <td colspan="2" class="left">
					  <select name="family_national" id="family_national" value="<%=family_national%>" style="width:90px">
				          <option value="" <% if family_rel = "" then %>selected<% end if %>>����</option>
				          <option value='������' <%If family_national = "������" then %>selected<% end if %>>������</option>
				          <option value='�ܱ���' <%If family_national = "�ܱ���" then %>selected<% end if %>>�ܱ���</option>
                      </select>
                      </td>
                    </tr>
                    <tr>
                      <th>�������(�ʼ�)</th>
                      <td colspan="2" class="left">
					  <input name="family_birthday" type="text" value="<%=family_birthday%>" style="width:70px;text-align:center" id="datepicker" readonly="true">
					  &nbsp;&nbsp;
					  <input type="radio" name="family_birthday_id" value="��" <% if family_birthday_id = "��" then %>checked<% end if %>>��
              		  <input name="family_birthday_id" type="radio" value="��" <% if family_birthday_id = "��" then %>checked<% end if %>>��
					  </td>
                      <th>�ֹε�Ϲ�ȣ</th>
                      <td colspan="2" class="left">
                      <input name="family_person1" type="text" id="family_person1" size="6" maxlength="6" value="<%=family_person1%>" >
					  -
                      <input name="family_person2" type="text" id="family_person2" size="7" maxlength="7" value="<%=family_person2%>" >
                      &nbsp;(���������ʼ�)
				      </td>
                   </tr>
                   <tr>
                      <th>����</th>
                      <td colspan="2" class="left">
					  <input name="family_job" type="text" id="family_job" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=family_job%>"></td>
                      <th>��ȭ��ȣ</th>
                      <td colspan="2" class="left">
                      <input name="family_tel_ddd" type="text" id="family_tel_ddd" size="3" maxlength="3" value="<%=family_tel_ddd%>" >
								  -
                      <input name="family_tel_no1" type="text" id="family_tel_no1" size="4" maxlength="4" value="<%=family_tel_no1%>" >
                                  -
                      <input name="family_tel_no2" type="text" id="family_tel_no2" size="4" maxlength="4" value="<%=family_tel_no2%>" >
					  </td>
                    </tr>
                    <tr>
                      <th>���ſ���</th>
                      <td colspan="2" class="left">
					  <input type="radio" name="family_live" value="����" <% if family_live = "����" then %>checked<% end if %>>���� 
              		  <input name="family_live" type="radio" value="����" <% if family_live = "����" then %>checked<% end if %>>����
                      <th>�ξ簡��</th>
                      <td colspan="2" class="left">
					  <input type="radio" name="family_support_yn" value="Y" <% if family_support_yn = "Y" then %>checked<% end if %>>�ξ� 
              		  <input name="family_support_yn" type="radio" value="N" <% if family_support_yn = "N" then %>checked<% end if %>>����
					  </td>
                    </tr>
                    <tr>
                      <th>�����</th>
                      <td colspan="2" class="left">
					  <input type="checkbox" name="disab_check" value="Y" <% if family_disab = "Y" then %>checked<% end if %> id="disab_check">�����
              		  <input type="checkbox" name="merit_check" value="Y" <% if family_merit = "Y" then %>checked<% end if %> id="merit_check">����������
                      <input type="checkbox" name="serius_check" value="Y" <% if family_serius = "Y" then %>checked<% end if %> id="serius_check">����ȯ��
					  </td>
                      <th>���α��ʻ�Ȱ����</th>
                      <td colspan="2" class="left">
					  <input type="checkbox" name="pensioner_check" value="Y" <% if family_pensioner = "Y" then %>checked<% end if %> id="pensioner_check">��
                    </tr>
                    <tr>
                      <th>�Ծ翩��</th>
                      <td class="left">
					  <input type="checkbox" name="holt_check" value="Y" <% if family_holt = "Y" then %>checked<% end if %> id="holt_check">��
                      </td>
                      <th>�Ծ�����</th>
                      <td class="left">
              		  <input name="family_holt_date" type="text" value="<%=family_holt_date%>" style="width:70px;text-align:center" id="datepicker1" readonly="true">
					  </td>
                      <th>��Ź�Ƶ�</th>
                      <td class="left">
					  <input type="checkbox" name="witak_check" value="Y" <% if family_witak = "Y" then %>checked<% end if %> id="witak_check">��
                    </tr>
                    <tr>
                      <th>�ڳ����</th>
                      <td colspan="5" class="left">
					  <input type="checkbox" name="children_check" value="Y" <% if family_children = "Y" then %>checked<% end if %> id="children_check">��&nbsp;&nbsp;(6���̸� �ڳ��ǰ�� �������� �߰����� üũ)
                      </td>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

