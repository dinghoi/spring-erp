<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
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
Dim f_seq, f_birthday_id, f_live, f_support_yn, f_national
Dim curr_date, title_line

f_birthday_id = "��"
f_live = "����"
f_support_yn = "N"
f_national = "������"

curr_date = Mid(CStr(Now()), 1, 10)
title_line = "�������� ���"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>ȸ������</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "" );
			});

			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.f_rel == ""){
					alert('���踦 �����ϼ���');
					frm.f_rel.focus();
					return false;
				}

				if(document.frm.f_name.value == ""){
					alert('������ �Է��ϼ���');
					frm.f_name.focus();
					return false;
				}

				if(document.frm.f_birthday.value == ""){
					alert('��������� �Է��ϼ���');
					frm.f_birthday.focus();
					return false;
				}

				if(document.frm.f_tel_ddd.value == ""){
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.family_tel_no1.focus();
					return false;
				}

				if(document.frm.f_tel_no1.value == ""){
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.family_tel_no1.focus();
					return false;
				}

				if(document.frm.f_tel_no2.value ==""){
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.family_tel_no2.focus();
					return false;
				}

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
			<form action="/member/member_family_proc.asp" method="post" name="frm">
			<div class="gView">
			  <table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="15%" >
					<col width="17%" >
					<col width="15%" >
					<col width="18%" >
					<col width="15%" >
					<col width="*" >
				</colgroup>
				<tbody>
				<tr>
					<th style="background:#FFFFE6">����</th>
					<td colspan="5" class="left" bgcolor="#FFFFE6">
						<input type="text" name="m_name" id="m_name" size="14" value="<%=m_name%>" class="no-input" readonly="true"/>
					</td>
				</tr>
				<tr>
				  <th>����<span style="color:red;">*</span></th>
				  <td colspan="5" class="left">
					  <select name="f_rel" id="f_rel" style="width:100px">
						  <option value="">����</option>
						  <option value='��'>��</option>
						  <option value='��'>��</option>
						  <option value='����'>����</option>
						  <option value='�Ƴ�'>�Ƴ�</option>
						  <option value='�Ƶ�'>�Ƶ�</option>
						  <option value='��'>��</option>
						  <option value='����'>����</option>
						  <option value='����'>����</option>
						  <option value='������'>������</option>
						  <option value='������'>������</option>
						  <option value='�ú�'>�ú�</option>
						  <option value='�ø�'>�ø�</option>
						  <option value='����'>����</option>
						  <option value='���'>���</option>
						  <option value='��(�����ڸ�)'>��(�����ڸ�)</option>
						  <option value='��(�����ڸ�)'>��(�����ڸ�)</option>
						  <option value='��(�����ڸ�)'>��(�����ڸ�)</option>
						  <option value='����'>����</option>
						  <option value='��(�����ڸ�)'>��(�����ڸ�)</option>
						  <option value='����'>����</option>
						  <option value='�ճ�'>�ճ�</option>
						  <option value='�ں�'>�ں�</option>
						  <option value='�պ�'>�պ�</option>
						  <option value='��Ÿ����'>��Ÿ����</option>
					  </select>
					  &nbsp;��Ź�Ƶ��ΰ��� ��Ÿ���踦 �����ϼ���.
				  </td>
				</tr>
				<tr>
				  <th>����<span style="color:red;">*</span></th>
				  <td colspan="2" class="left">
					<input type="text" name="f_name" id="f_name" size="14"/></td>
				  <th>�����</th>
				  <td colspan="2" class="left">
					  <select name="f_national" id="f_national" style="width:90px">
						  <option value="">����</option>
						  <option value='������'>������</option>
						  <option value='�ܱ���'>�ܱ���</option>
					  </select>
				  </td>
				</tr>
				<tr>
					<th>�������<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="f_birthday" id="datepicker" style="width:70px;text-align:center" readonly="true"/>
						&nbsp;&nbsp;
						<input type="radio" name="f_birthday_id" id="f_birthday_id" value="��" checked/>��
						<input type="radio" name="f_birthday_id" id="f_birthday_id" value="��"/>��
					</td>
					<th>�ֹε�Ϲ�ȣ</th>
					<td colspan="2" class="left">
						<input type="text" name="f_person1" id="f_person1" size="6" maxlength="6" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						-
						<input type="text" name="f_person2" id="f_person2" size="7" maxlength="7" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						(���������ʼ�)
					</td>
			   </tr>
			   <tr>
					<th>����</th>
					<td colspan="2" class="left">
						<input type="text" name="f_job" id="f_job" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);"/>
					</td>
					<th>��ȭ��ȣ<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="f_tel_ddd" id="f_tel_ddd" size="3" maxlength="3" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="f_tel_no1" id="f_tel_no1" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="f_tel_no2" id="f_tel_no2" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
					</td>
				</tr>
				<tr>
					<th>���ſ���</th>
					<td colspan="2" class="left">
						<input type="radio" name="f_live" id="f_live" value="����" checked/>����
						<input type="radio" name="f_live" id="f_live" value="����"/>����
					<th>�ξ簡��</th>
					<td colspan="2" class="left">
						<input type="radio" name="f_support_yn" id="f_support_yn" value="Y" checked/>�ξ�
						<input type="radio" name="f_support_yn" id="f_support_yn" value="N"/>����
				  </td>
				</tr>
				<tr>
					<th>�����</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="disab_check" id="disab_check" value="Y"/>�����
						<input type="checkbox" name="merit_check" id="merit_check" value="Y"/>����������
						<input type="checkbox" name="serius_check" id="serius_check" value="Y"/>����ȯ��
					</td>
					<th>���α��ʻ�Ȱ����</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="pensioner_check" id="pensioner_check" value="Y"/>��
					</td>
				</tr>
				<tr>
					<th>�Ծ翩��</th>
					<td class="left">
						<input type="checkbox" name="holt_check" id="holt_check" value="Y"/>��
					</td>
					<th>�Ծ�����</th>
					<td class="left">
						<input name="f_holt_date" type="text" id="datepicker1" style="width:70px;text-align:center" readonly="true"/>
					</td>
					<th>��Ź�Ƶ�</th>
					<td class="left">
						<input type="checkbox" name="witak_check" id="witak_check" value="Y"/>��
					</td>
				</tr>
				<tr>
					<th>�ڳ����</th>
					<td colspan="5" class="left">
						<input type="checkbox" name="children_check" id="children_check" value="Y"/>��&nbsp;&nbsp;(6���̸� �ڳ��ǰ�� �������� �߰����� üũ)
					</td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
			</div>
			</form>
		</div>
	</body>
</html>