<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim u_type, family_empno, family_seq, emp_name
Dim family_rel, family_name, family_birthday, family_birthday_id
Dim family_job, family_live, family_person1, family_person2
Dim family_tel_ddd, family_tel_no1, family_tel_no2, family_support_yn
Dim family_national, family_disab, family_merit, family_serius
Dim family_pensioner, family_witak, family_holt, family_holt_date, family_children
Dim curr_date, title_line, rsFamily

u_type = Request.QueryString("u_type")
family_empno = Request.QueryString("family_empno")
family_seq = Request.QueryString("family_seq")
emp_name = Request.QueryString("emp_name")

family_rel = ""
family_name = ""
family_birthday = ""
family_birthday_id = "��"
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

curr_date = Mid(CStr(Now()), 1, 10)
title_line = "�������� ���"

If u_type = "U" Then
	objBuilder.Append "SELECT family_rel, family_name, family_birthday, family_birthday_id, family_job, "
	objBuilder.Append "	family_live, family_person1, family_person2, family_tel_ddd, family_tel_no1, "
	objBuilder.Append "	family_tel_no2, family_support_yn, family_national, family_disab, family_merit, "
	objBuilder.Append "	family_serius, family_pensioner, family_witak, family_holt, family_holt_date, "
	objBuilder.Append "	family_children "
	objBuilder.Append "FROM emp_family "
	objBuilder.Append "WHERE family_empno = '"&family_empno&"' AND family_seq = '"&family_seq&"';"

	Set rsFamily = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	family_rel = rsFamily("family_rel")
    family_name = rsFamily("family_name")
    family_birthday = rsFamily("family_birthday")
    family_birthday_id = rsFamily("family_birthday_id")
    family_job = rsFamily("family_job")
    family_live = rsFamily("family_live")
    family_person1 = rsFamily("family_person1")
    family_person2 = rsFamily("family_person2")
	family_tel_ddd = rsFamily("family_tel_ddd")
    family_tel_no1 = rsFamily("family_tel_no1")
    family_tel_no2 = rsFamily("family_tel_no2")
	family_support_yn = rsFamily("family_support_yn")
	family_national = rsFamily("family_national")
    family_disab = rsFamily("family_disab")
	family_merit = rsFamily("family_merit")
    family_serius = rsFamily("family_serius")
    family_pensioner = rsFamily("family_pensioner")
    family_witak = rsFamily("family_witak")
    family_holt = rsFamily("family_holt")
    family_holt_date = rsFamily("family_holt_date")
	family_children = rsFamily("family_children")

	If family_birthday = "1900-01-01"  Then
	   family_birthday = ""
	end If

	If family_holt_date = "1900-01-01"  Then
	   family_holt_date = ""
	End If

	rsFamily.Close() : Set rsFamily = Nothing

	title_line = "�������� ����"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ�������</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			//�������
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
			});

			//�Ծ�����
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=family_holt_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.family_rel == ""){
					alert('���踦 �������ּ���.');
					frm.family_rel.focus();
					return false;
				}

				if(document.frm.family_name.value == ""){
					alert('������ �Է����ּ���.');
					frm.family_name.focus();
					return false;
				}

				if(document.frm.family_birthday.value == ""){
					alert('��������� �Է����ּ���.');
					frm.family_birthday.focus();
					return false;
				}

				if(document.frm.family_tel_ddd.value == ""){
					alert('�޴�����ȣ�� �Է����ּ���.');
					frm.family_tel_ddd.focus();
					return false;
				}

				if(document.frm.family_tel_no1.value == ""){
					alert('�޴�����ȣ�� �Է����ּ���.');
					frm.family_tel_no1.focus();
					return false;
				}

				if(document.frm.family_tel_no2.value ==""){
					alert('�޴�����ȣ�� �Է����ּ���.');
					frm.family_tel_no2.focus();
					return false;
				}

				/*if(document.frm.family_support_yn.value == ""){
					alert('�ξ簡�� ���θ� �������ּ���.');
					frm.family_support_yn.focus();
					return false;
				}*/

				var result = confirm('��� �Ͻðڽ��ϱ�?');

				if(result){
					return true;
				}else{
					return false
				};
			}
        </script>

		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body>
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/person/insa_family_add_save.asp" method="post" name="frm">
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
						<input type="text" name="family_empno" id="family_empno" size="14" value="<%=family_empno%>" readonly class="no-input"/>
						<input type="hidden" name="family_seq" value="<%=family_seq%>"/>
					</td>
					<th style="background:#FFFFE6">����</th>
					<td colspan="3" class="left" bgcolor="#FFFFE6">
						<input type="text" name="emp_name" id="emp_name" size="14" value="<%=emp_name%>" readonly class="no-input"/>
					</td>
				</tr>
				<tr>
				  <th>����<span style="color:red;">*</span></th>
				  <td colspan="5" class="left">
					  <select name="family_rel" id="family_rel" value="<%=family_rel%>" style="width:100px;">
						  <option value="" <%If family_rel = "" Then %>selected<%End If %>>����</option>
						  <option value='��' <%If family_rel = "��" Then %>selected<%End If %>>��</option>
						  <option value='��' <%If family_rel = "��" Then %>selected<%End If %>>��</option>
						  <option value='����' <%If family_rel = "����" Then %>selected<%End If %>>����</option>
						  <option value='�Ƴ�' <%If family_rel = "�Ƴ�" Then %>selected<%End If %>>�Ƴ�</option>
						  <option value='�Ƶ�' <%If family_rel = "�Ƶ�" Then %>selected<%End If %>>�Ƶ�</option>
						  <option value='��' <%If family_rel = "��" Then %>selected<%End If %>>��</option>
						  <option value='����' <%If family_rel = "����" Then %>selected<%End If %>>����</option>
						  <option value='����' <%If family_rel = "����" Then %>selected<%End If %>>����</option>
						  <option value='������' <%If family_rel = "������" Then %>selected<%End If %>>������</option>
						  <option value='������' <%If family_rel = "������" Then %>selected<%End If %>>������</option>
						  <option value='�ú�' <%If family_rel = "�ú�" Then %>selected<%End If %>>�ú�</option>
						  <option value='�ø�' <%If family_rel = "�ø�" Then %>selected<%End If %>>�ø�</option>
						  <option value='����' <%If family_rel = "����" Then %>selected<%End If %>>����</option>
						  <option value='���' <%If family_rel = "���" Then %>selected<%End If %>>���</option>
						  <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" Then %>selected<%End If %>>��(�����ڸ�)</option>
						  <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" Then %>selected<%End If %>>��(�����ڸ�)</option>
						  <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" Then %>selected<%End If %>>��(�����ڸ�)</option>
						  <option value='����' <%If family_rel = "����" Then %>selected<%End If %>>����</option>
						  <option value='��(�����ڸ�)' <%If family_rel = "��(�����ڸ�)" Then %>selected<%End If %>>��(�����ڸ�)</option>
						  <option value='����' <%If family_rel = "����" Then %>selected<%End If %>>����</option>
						  <option value='�ճ�' <%If family_rel = "�ճ�" Then %>selected<%End If %>>�ճ�</option>
						  <option value='�ں�' <%If family_rel = "�ں�" Then %>selected<%End If %>>�ں�</option>
						  <option value='�պ�' <%If family_rel = "�պ�" Then %>selected<%End If %>>�պ�</option>
						  <option value='��Ÿ����' <%If family_rel = "��Ÿ����" Then %>selected<%End If %>>��Ÿ����</option>
					  </select>
					  &nbsp;
					  (<span style="color:red;font-size:11px;">��Ź�Ƶ��ΰ��� ��Ÿ���踦 �����Ͻʽÿ�.</span>)
				  </td>
				</tr>
				<tr>
				  <th>����<span style="color:red;">*</span></th>
				  <td colspan="2" class="left">
					<input type="text" name="family_name" id="family_name" size="14" value="<%=family_name%>"/></td>
				  <th>�����</th>
				  <td colspan="2" class="left">
					  <select name="family_national" id="family_national" value="<%=family_national%>" style="width:90px">
						  <option value="" <%If family_rel = "" Then %>selected<%End If %>>����</option>
						  <option value='������' <%If family_national = "������" Then %>selected<%End If %>>������</option>
						  <option value='�ܱ���' <%If family_national = "�ܱ���" Then %>selected<%End If %>>�ܱ���</option>
					  </select>
				  </td>
				</tr>
				<tr>
					<th>�������<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="family_birthday" value="<%=family_birthday%>" style="width:70px;text-align:center" id="datepicker" readonly="true"/>
						&nbsp;&nbsp;
						<input type="radio" name="family_birthday_id" value="��" <%If family_birthday_id = "��" Then %>checked<%End If %>/>��
						<input type="radio" name="family_birthday_id" value="��" <%If family_birthday_id = "��" Then %>checked<%End If %>/>��
					</td>
					<th>�ֹε�Ϲ�ȣ</th>
					<td colspan="2" class="left">
						<input type="text" name="family_person1" id="family_person1" style="width:40px;" maxlength="6" value="<%=family_person1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						-
						<input type="text" name="family_person2" id="family_person2" style="width:50px;"  maxlength="7" value="<%=family_person2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						&nbsp;
						(<span style="color:red;font-size:11px;">�������� �� �ʼ�</span>)
					</td>
			   </tr>
			   <tr>
					<th>����</th>
					<td colspan="2" class="left">
						<input name="family_job" type="text" id="family_job" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=family_job%>"/>
					</td>
					<th>�޴�����ȣ<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="family_tel_ddd" id="family_tel_ddd" size="3" maxlength="3" value="<%=family_tel_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="family_tel_no1" id="family_tel_no1" size="4" maxlength="4" value="<%=family_tel_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="family_tel_no2" id="family_tel_no2" size="4" maxlength="4" value="<%=family_tel_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
					</td>
				</tr>
				<tr>
					<th>���ſ���</th>
					<td colspan="2" class="left">
						<input type="radio" name="family_live" value="����" <%If family_live = "����" Then %>checked<%End If %>/>����
						<input type="radio" name="family_live" value="����" <%If family_live = "����" Then %>checked<%End If %>/>����
					<th>�ξ簡��</th>
					<td colspan="2" class="left">
						<input type="radio" name="family_support_yn" value="Y" <%If family_support_yn = "Y" Then %>checked<%End If %>/>�ξ�
						<input type="radio" name="family_support_yn" value="N" <%If family_support_yn = "N" Then %>checked<%End If %>/>����
				  </td>
				</tr>
				<tr>
					<th>�����</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="disab_check" value="Y" <%If family_disab = "Y" Then %>checked<%End If %> id="disab_check"/>�����
						<input type="checkbox" name="merit_check" value="Y" <%If family_merit = "Y" Then %>checked<%End If %> id="merit_check"/>����������
						<input type="checkbox" name="serius_check" value="Y" <%If family_serius = "Y" Then %>checked<%End If %> id="serius_check"/>����ȯ��
					</td>
					<th>���α��ʻ�Ȱ����</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="pensioner_check" value="Y" <%If family_pensioner = "Y" Then %>checked<%End If %> id="pensioner_check"/>��
					</td>
				</tr>
				<tr>
					<th>�Ծ翩��</th>
					<td class="left">
						<input type="checkbox" name="holt_check" value="Y" <%If family_holt = "Y" Then %>checked<%End If %> id="holt_check"/>��
					</td>
					<th>�Ծ�����</th>
					<td class="left">
						<input type="text" name="family_holt_date" value="<%=family_holt_date%>" style="width:70px;text-align:center" id="datepicker1" readonly/>
					</td>
					<th>��Ź�Ƶ�</th>
					<td class="left">
						<input type="checkbox" name="witak_check" value="Y" <%If family_witak = "Y" Then %>checked<%End If %> id="witak_check"/>��
					</td>
				</tr>
				<tr>
					<th>�ڳ����</th>
					<td colspan="5" class="left">
						<input type="checkbox" name="children_check" value="Y" <%If family_children = "Y" Then %>checked<%End If %> id="children_check"/>��&nbsp;&nbsp;
						(<span style="color:red;font-size:11px;">6���̸� �ڳ��ǰ�� �������� �߰����� üũ</span>)
					</td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="<%If u_type = "U" Then%>����<%Else%>���<%End If%>" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>"/>
			</form>
		</div>
	</body>
</html>