<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim emp_reg_date, emp_reg_user, emp_disabled, emp_pay_id
Dim emp_mod_date, title_line
Dim rsInfo, arrInfo, i
Dim m_ename, m_birthday, m_birthday_id, m_person1, m_person2
Dim m_sex, m_tel_ddd, m_tel_no1, m_tel_no2, m_hp_ddd, m_hp_no1, m_hp_no2
Dim m_zipcode, m_sido, m_gugun, m_dong, m_addr, m_emergency_tel, m_sawo_id
Dim m_hobby, m_disabled, m_disab_grade, m_military_id, m_military_grade
Dim m_military_date1, m_military_date2, m_military_comm, m_marry_date
Dim m_faith, m_last_edu, m_image, m_reg_date, photo_image, att_file

m_seq = Request.QueryString("m_seq")

title_line = "ä�� ����"

emp_reg_date = Now()
emp_reg_user = user_name

emp_disabled = "�ش���׾���"
emp_pay_id = "0"
mg_group = "1"
emp_mod_date = ""

'ȸ������ ���� ��ȸ
objBuilder.Append "SELECT m_name, m_ename, m_birthday, m_birthday_id, m_person1, m_person2, "
objBuilder.Append "	m_sex, m_tel_ddd, m_tel_no1, m_tel_no2, m_hp_ddd, m_hp_no1, m_hp_no2, "
objBuilder.Append "	m_zipcode, m_sido, m_gugun, m_dong, m_addr, m_emergency_tel, m_sawo_id, "
objBuilder.Append "	m_hobby, m_disabled, m_disab_grade, m_military_id, m_military_grade, "
objBuilder.Append "	m_military_date1, m_military_date2, m_military_comm, m_marry_date, "
objBuilder.Append "	m_faith, m_last_edu, m_image, m_reg_date "
objBuilder.Append "FROM member_info "
objBuilder.Append "WHERE m_seq = '"&m_seq&"' "

Set rsInfo = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsInfo.EOF Then
	arrInfo = rsInfo.getRows()
End If
rsInfo.Close() : Set rsInfo = Nothing

If IsArray(arrInfo) Then
	For i = LBound(arrInfo) To UBound(arrInfo, 2)
		m_name = arrInfo(0, i)
		m_ename = arrInfo(1, i)
		m_birthday = arrInfo(2, i)
		m_birthday_id = arrInfo(3, i)
		m_person1 = arrInfo(4, i)
		m_person2 = arrInfo(5, i)
		m_sex = arrInfo(6, i)
		m_tel_ddd = arrInfo(7, i)
		m_tel_no1 = arrInfo(8, i)
		m_tel_no2 = arrInfo(9, i)
		m_hp_ddd = arrInfo(10, i)
		m_hp_no1 = arrInfo(11, i)
		m_hp_no2 = arrInfo(12, i)
		m_zipcode = arrInfo(13, i)
		m_sido = arrInfo(14, i)
		m_gugun = arrInfo(15, i)
		m_dong = arrInfo(16, i)
		m_addr = arrInfo(17, i)
		m_emergency_tel = arrInfo(18, i)
		m_sawo_id = arrInfo(19, i)
		m_hobby = arrInfo(20, i)
		m_disabled = arrInfo(21, i)
		m_disab_grade = arrInfo(22, i)
		m_military_id = arrInfo(23, i)
		m_military_grade = arrInfo(24, i)
		m_military_date1 = arrInfo(25, i)
		m_military_date2 = arrInfo(26, i)
		m_military_comm = arrInfo(27, i)
		m_marry_date = arrInfo(28, i)
		m_faith = arrInfo(29, i)
		m_last_edu = arrInfo(30, i)
		m_image = arrInfo(31, i)
		m_reg_date = arrInfo(32, i)
	Next

	If m_military_date1 = "1900-01-01" Then
		m_military_date1 = ""
	End If

	If m_military_date2 = "1900-01-01" Then
		m_military_date2 = ""
	End If

	If m_marry_date = "1900-01-01" Then
		m_marry_date = ""
	End If
Else
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('����ġ ���� ������ �߻��߽��ϴ�.');"
	Response.Write "	window.close();"
	Response.Write "</script>"
	Response.End
End If

If f_toString(m_image, "") <> "" Then
	photo_image = "/emp_photo/"&m_image
	att_file = m_image
Else
	photo_image = ""
	att_file = ""
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}

			//�����Ի���
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "" );
			});

			//�Ի���
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "" );
			});

			//���������
			$(function(){
				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "" );
			});

			//�ټӱ����
			$(function(){
				$( "#datepicker3" ).datepicker();
				$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker3" ).datepicker("setDate", "" );
			});

			//���������
			$(function(){
				$( "#datepicker4" ).datepicker();
				$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker4" ).datepicker("setDate", "" );
			});

			//����������
			$(function(){
				$( "#datepicker6" ).datepicker();
				$( "#datepicker6" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker6" ).datepicker("setDate", "" );
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
				if($('#emp_no').val() === ''){
					alert('�����ȣ�� �Է��� �ּ���.');
					$('#emp_no').focus();
					return false;
				}

				if($('#emp_org_code').val() === ''){
					alert('�Ҽ��� ������ �ּ���.');
					$('#emp_org_code').focus();
					return false;
				}

				if($('#emp_type').val() === ''){
					alert('���������� ������ �ּ���.');
					$('#emp_type').focus();
					return false;
				}

				if($('#emp_grade').val() === ''){
					alert('������ ������ �ּ���.');
					$('#emp_grade').focus();
					return false;
				}

				if($('#emp_job').val() === ''){
					alert('������ ������ �ּ���.');
					$('#emp_job').focus();
					return false;
				}

				if($('#emp_position').val() === ''){
					alert('��å�� ������ �ּ���.');
					$('#emp_position').focus();
					return false;
				}

				if($('#emp_jikmu').val() === ''){
					alert('������ ������ �ּ���.');
					$('#emp_jikmu').focus();
					return false;
				}

				if($('#datepicker').val() === ''){
					alert('�����Ի����� �Է��� �ּ���.');
					$('#datepicker').focus();
					return false;
				}

				if($('#datepicker1').val() === ''){
					alert('�Ի����� �Է��� �ּ���.');
					$('#datepicker1').focus();
					return false;
				}

				if($('#datepicker2').val() === ''){
					alert('����������� �Է��ϼ���.');
					$('#datepicker2').focus();
					return false;
				}

				if($('#datepicker3').val() === ''){
					alert('�ټӱ������ �Է��� �ּ���.');
					$('#datepicker3').focus();
					return false;
				}

				if($('#datepicker4').val() === ''){
					alert('����������� �Է��� �ּ���.');
					$('#datepicker4').focus();
					return false;
				}

				if($('#emp_first_date').val() > $('#emp_in_date').val()){
					alert('�����Ի����� �Ի��Ϻ��� �ʽ��ϴ�.');
					$('#emp_first_date').focus();
					return false;
				}

				if($('#emp_in_date').val() > $('#emp_end_gisan').val()){
					alert('����������� �Ի��Ϻ��� �����ϴ�.');
					$('#emp_end_gisan').focus();
					return false;
				}

				if($('#emp_in_date').val() > $('#emp_yuncha_date').val()){
					alert('����������� �Ի��Ϻ��� �����ϴ�.');
					$('#emp_yuncha_date').focus();
					return false;
				}

				if($('#emp_email').val() === ''){
					alert('�̸��� �ּҸ� �Է��� �ּ���.');
					$('#emp_email').focus();
					return false;
				}

				if($('#cost_center').val() === ''){
					alert('��뱸���� ������ �ּ���.');
					$('#cost_center').focus();
					return false;
				}
				/*
				if($('#mg_group').val() === ''){
					alert('�����׷쿩�θ� üũ�� �ּ���.');
					$('#mg_group').focus();
					return false;
				}*/

				if($('#emp_pay_id').val() === ''){
					alert('�޿������ ������ �ּ���.');
					$('#emp_pay_id').focus();
					return false;
				}

				if($('#dz_id').val() === ''){
					alert('�޿�ID�� �Է��� �ּ���.');
					$('#dz_id').focus();
					return false;
				}

				if($('#cost_center').val() === '����������'){
					if($('#emp_reside_company').val() === '') {
						alert('����óȸ�縦 ������ �ּ���.');
						$('#emp_reside_company').focus();
						return false;
					}
				}

				var result = confirm('���� ó���Ͻðڽ��ϱ�?');

				if(result == true){
					return true;
				}
				return false;
			}
			/*
			function file_browse(){
           		document.frm.att_file.click();
           		document.frm.text1.value=document.frm.att_file.value;
			}*/
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_approve_proc.asp" method="post" name="frm">
					<input type="hidden" name="m_seq" id="m_seq" value="<%=m_seq%>">
					<input type="hidden" name="emp_end_date" id="emp_end_date"/>
					<input type="hidden" name="emp_org_baldate" id="emp_org_baldate"/>
					<input type="hidden" name="emp_grade_date" id="emp_grade_date"/>
					<input type="hidden" name="emp_image" id="emp_image" value="<%=m_image%>"/>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
						</colgroup>
						<tbody>
							<tr>
								<td colspan="2" rowspan="4" class="left">
									<img src="<%=photo_image%>" width="110" height="120" alt="�������"/>
								</td>
								<th>��� ��ȣ<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_no" id="emp_no" size="9" maxlength="6" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
								<th>����(�ѱ�)</th>
								<td class="left">
									<input type="text" name="emp_name" id="emp_name" value="<%=m_name%>" size="9" readonly class="no-input"/>
								</td>
								<th>����(����)</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_ename" id="emp_ename" value="<%=m_ename%>" style="width:160px;" maxlength="20" readonly class="no-input"/>
								</td>
								<th>�������</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_birthday" size="10" id="emp_birthday" style="width:70px;" value="<%=m_birthday%>" readonly class="no-input"/>
									&nbsp;��&nbsp;
									<input type="text" name="emp_birthday_id" value="<%=m_birthday_id%>" style="width:20px;text-align:center;" readonly class="no-input"/>
								</td>
							</tr>
							<tr>
								<th>�Ҽ�<span style="color:red;">*</span></th>
								<td colspan="3" class="left">
									<input type="text" name="emp_org_code" id="emp_org_code" style="width:40px" readonly/>
									&nbsp;��&nbsp;
									<input type="text" name="emp_org_name" id="emp_org_name" style="width:120px" readonly/>

									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=org','�Ҽ� ��ȸ','scrollbars=yes,width=800,height=400')">����</a>
								</td>
								<th>����<span style="color:red;">*</span></th>
								<td colspan="5" class="left">
									<input type="text" name="emp_company" id="emp_company" style="width:100px" readonly/>
									<input type="text" name="emp_bonbu" id="emp_bonbu" style="width:120px" readonly/>
									<input type="text" name="emp_saupbu" id="emp_saupbu" style="width:120px" readonly/>
									<input type="text" name="emp_team" id="emp_team" style="width:120px" readonly/>

									<input type="hidden" name="emp_reside_place" id="emp_reside_place"/>
									<input type="hidden" name="emp_org_level" id="emp_org_level"/>
								</td>
							</tr>
            				<tr>
            					<th>��������<span style="color:red;">*</span></th>
            					<td class="left">
            						<select name="emp_type" id="emp_type" style="width:90px">
			            				<option value="">����</option>
										<option value='����'>����</option>
										<option value='����'>����</option>
										<option value='�����'>�����</option>
					                </select>
								</td>
								<th>����<span style="color:red;">*</span></th>
								<td class="left">
                				<%
								Dim rsGrade, rsJob, rsPosition

								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '02' ORDER BY emp_etc_code ASC;"

                				Set rsGrade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="emp_grade" id="emp_grade" style="width:90px">
                  						<option value="">����</option>
								<%
								Do Until rsGrade.EOF
								%>
                						<option value='<%=rsGrade("emp_etc_name")%>'><%=rsGrade("emp_etc_name")%></option>
                				<%
                					rsGrade.MoveNext()
                				Loop
                				rsGrade.Close() : Set rsGrade = Nothing
							  	%>
            						</select>
								</td>
								<th>����<span style="color:red;">*</span></th>
								<td class="left">
								<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '03' ORDER BY emp_etc_code ASC "

								Set rsJob = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="emp_job" id="emp_job" style="width:90px">
                  						<option value="">����</option>
                				<%
                				Do Until rsJob.EOF
			  				  	%>
                						<option value='<%=rsJob("emp_etc_name")%>'><%=rsJob("emp_etc_name")%></option>
                				<%
									rsJob.MoveNext()
								Loop
								rsJob.Close() : Set rsJob = Nothing
							  	%>
            						</select>
								</td>
								<th>��å<span style="color:red;">*</span></th>
								<td class="left">
                				<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '04' ORDER BY emp_etc_code ASC;"

								Set rsPosition = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="emp_position" id="emp_position" style="width:90px">
                  						<option value="">����</option>
                				<%
                				Do Until rsPosition.EOF
			  				  	%>
                					<option value='<%=rsPosition("emp_etc_name")%>'><%=rsPosition("emp_etc_name")%></option>
                				<%
									rsPosition.MoveNext()
								Loop
								rsPosition.Close() : Set rsPosition = Nothing
							 	%>
            						</select>
								</td>
								<th>����<span style="color:red;">*</span></th>
								<td class="left">
                				<%
								Dim rsJikmu

								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '05' ORDER BY emp_etc_code ASC;"

								Set rsJikmu = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="emp_jikmu" id="emp_jikmu" style="width:90px">
                  						<option>����</option>
                				<%
								Do Until rsJikmu.EOF
			  				  	%>
                						<option value='<%=rsJikmu("emp_etc_name")%>'><%=rsJikmu("emp_etc_name")%></option>
                				<%
									rsJikmu.MoveNext()
								Loop
								rsJikmu.Close() : Set rsJikmu = Nothing
							  	%>
            						</select>
								</td>
							</tr>
							<tr>
								<th>�����Ի���<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_first_date" size="10" id="datepicker" style="width:70px;" readonly="true"/>&nbsp;
								</td>
								<th>�Ի���<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_in_date" size="10" id="datepicker1" style="width:70px;" readonly="true"/>&nbsp;
								</td>
								<th>���������<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_end_gisan" size="10" id="datepicker2" style="width:70px;" readonly="true"/>
								</td>
								<th>�ټӱ����<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_gunsok_date" size="10" id="datepicker3" style="width:70px;" readonly="true"/>
								</td>
								<th>���������<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_yuncha_date" size="10" id="datepicker4" style="width:70px;" readonly="true"/>
								</td>
							</tr>
							<tr>
              					<th colspan="2">�ֹι�ȣ</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_person1" id="emp_person1" size="4" maxlength="6" value="<%=m_person1%>" readonly class="no-input"/>
									��
									<input type="text" name="emp_person2" id="emp_person2" size="5" maxlength="7" value="<%=m_person2%>" readonly class="no-input"/>
									<input type="text" name="emp_sex" id="emp_sex" value="<%=m_sex%>" style="width:20px;text-align:center;" readonly class="no-input"/>
								</td>
              					<th>��ȭ��ȣ</th>
								<td colspan="3" class="left">
									<input type="text" name="emp_tel_ddd" id="emp_tel_ddd" size="3" maxlength="3" value="<%=m_tel_ddd%>" readonly class="no-input"/>
									-
									<input type="text" name="emp_tel_no1" id="emp_tel_no1" size="4" maxlength="4" value="<%=m_tel_no1%>" readonly class="no-input"/>
									-
									<input type="text" name="emp_tel_no2" id="emp_tel_no2" size="4" maxlength="4" value="<%=m_tel_no2%>" readonly class="no-input"/>
								</td>
								<th>�޴�����ȣ</th>
								<td colspan="3" class="left">
									<input type="text" name="emp_hp_ddd" id="emp_hp_ddd" size="3" maxlength="3" value="<%=m_hp_ddd%>" readonly class="no-input"/>
									-
									<input type="text" name="emp_hp_no1" id="emp_hp_no1" size="4" maxlength="4" value="<%=m_hp_no1%>" readonly class="no-input"/>
									-
									<input type="text" name="emp_hp_no2" id="emp_hp_no2" size="4" maxlength="4" value="<%=m_hp_no2%>" readonly class="no-input"/>
								</td>
							</tr>
							<tr>
								<th colspan="2">�ּ�(��)</th>
								<td colspan="7" class="left">
									<input type="text" name="emp_zipcode" id="emp_zipcode" style="width:50px;" value="<%=m_zipcode%>" class="no-input" readonly/>
									-
									<input type="text" name="emp_sido" id="emp_sido" style="width:100px" value="<%=m_sido%>" readonly class="no-input"/>
									<input type="text" name="emp_gugun" id="emp_gugun" style="width:150px" value="<%=m_gugun%>" readonly class="no-input"/>
									<input type="text" name="emp_dong" id="emp_dong" style="width:150px" value="<%=m_dong%>" readonly class="no-input"/>
									<input type="text" name="emp_addr" id="emp_addr" style="width:200px" value="<%=m_addr%>" readonly class="no-input"/>
              						<!--<a href="#" class="btnType03" onClick="pop_Window('/insa/jusoPopup.asp?gubun=juso','family_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>-->
								</td>
								<th>�̸��� �ּ�<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="emp_email" id="emp_email" size="12" /> @k-one.co.kr
								</td>
							</tr>
							<tr>
								<th colspan="2" class="first">�������Կ���</th>
								<td class="left">
									<input type="radio" name="emp_sawo_id" value="Y" <%If m_sawo_id = "Y" Then %>checked<%End If %>/>����
              						<input type="radio" name="emp_sawo_id" value="N" <%If m_sawo_id = "N" Then %>checked<%End If %>/>����
								</td>
								<th>����������</th>
                                <td class="left">
									<input type="text" name="emp_sawo_date" size="10" id="datepicker6" style="width:70px;"/>
                                </td>
								<th>��ȥ�����</th>
								<td class="left">
									<input type="text" name="emp_marry_date" size="10" style="width:70px;" value="<%=m_marry_date%>" readonly class="no-input"/>
								</td>
								<th>���</th>
								<td class="left">
									<input type="text" name="emp_hobby" id="emp_hobby" size="9" value="<%=m_hobby%>" readonly class="no-input"/>
								</td>
								<th>���/���</th>
								<td colspan="2" class="left">
            						<input type="text" name="emp_disabled" id="emp_disabled" style="width:90px;text-align:center;" value="<%=m_disabled%>" readonly class="no-input"/>
									-
									<input type="text" name="emp_disab_grade" id="emp_disab_grade" value="<%=m_disab_grade%>" style="width:30px" readonly class="no-input"/>
								</td>
							</tr>
							<tr>
								<th colspan="2" >��������</th>
								<td class="left">
									<input type="text" name="emp_military_id" id="emp_military_id" value="<%=m_military_id%>" style="width:30px" readonly class="no-input"/>
								</td>
								<th>�������</th>
								<td class="left">
									<input type="text" name="emp_military_grade" id="emp_military_grade" value="<%=m_military_grade%>" style="width:30px" readonly class="no-input"/>
								</td>
								<th>���� �����Ⱓ</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_military_date1" id="emp_military_date1" style="width:70px;" value="<%=m_military_date1%>" readonly class="no-input"/>
									��
									<input type="text" name="emp_military_date2" id="emp_military_date2" style="width:70px;" value="<%=m_military_date2%>" readonly class="no-input"/>
								</td>
								<th>��������</th>
								<td class="left">
									<input type="text" name="emp_military_comm" id="emp_military_comm" size="9" value="<%=m_military_comm%>" readonly class="no-input"/></td>
								</td>
								<th>����</th>
								<td class="left">
									<input type="text" name="emp_faith" id="emp_faith" style="width:50px;text-align:center;" value="<%=m_faith%>" readonly class="no-input"/>
								</td>
							</tr>
							<tr>
								<th colspan="2" class="first">�Ǳٹ���/�ּ�</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_stay_name" id="emp_stay_name" size="10"/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_stay_select.asp?gubun=stay','stayselect','scrollbars=yes,width=1000,height=400')">����</a>
								</td>
								<td colspan="4" class="left">
									<input type="hidden" name="emp_stay_code" id="emp_stay_code" readonly/>
									<input type="text" name="stay_sido" id="stay_sido" style="width:60px;" readonly/>
									<input type="text" name="stay_gugun" id="stay_gugun" style="width:60px;" readonly/>
									<input type="text" name="stay_dong" id="stay_dong" style="width:60px;" readonly/>
									<input type="text" name="stay_addr" id="stay_addr" style="width:190px;" readonly/>
								</td>
								<th>���׷�</th>
								<td class="left">
                					<input type="text" name="cost_group" id="cost_group" style="width:80px;" readonly/>
            					</td>
								<th>��󿬶�</th>
								<td class="left">
									<input type="text" name="emp_emergency_tel" id="emp_emergency_tel" size="10" value="<%=m_emergency_tel%>" readonly class="no-input"/>
								</td>
							</tr>
							<tr>
								<th colspan="2" class="first">������ȣ</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_extension_no" id="emp_extension_no" size="16" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
								<th>�����з�</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_last_edu" id="emp_last_edu" style="width:100px;text-align:center;" value="<%=m_last_edu%>" readonly class="no-input"/>
								</td>
								<th>��뱸��<span style="color:red;">*</span></th>
								<td class="left">
                				<%
								Dim rsCostType

                				objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC;"

								Set rsCostType = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="cost_center" id="cost_center" style="width:90px">
                  						<option value="">����</option>
                				<%
                				Do Until rsCostType.EOF
			  				  	%>
                						<option value='<%=rsCostType("emp_etc_name")%>'><%=rsCostType("emp_etc_name")%></option>
                				<%
									rsCostType.MoveNext()
								Loop
								rsCostType.Close() : Set rsCostType = Nothing
								DBConn.Close() : Set DBConn = Nothing
							  	%>
                					</select>
								</td>
								<th>�����׷쿩��</th>
								<td colspan="2" class="left">
									<input type="radio" name="mg_group" value="1" checked>�Ϲݱ׷�
									<input type="radio" name="mg_group" value="2">�����׷�
								</td>
							</tr>
							<tr>
								<th colspan="2" class="first">���� �����</th>
								<td colspan="2" class="left"><%=emp_reg_date%>&nbsp;(<%=emp_reg_user%>)</td>
								<th>����ó ȸ��</th>
								<td colspan="2" class="left">
									<input name="emp_reside_company" type="text" id="emp_reside_company" style="width:90px;" readonly>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_trade_search.asp?gubun=5','tradesearch','scrollbars=yes,width=600,height=400')">ã��</a>
								</td>
								<th>�޿����<span style="color:red;">*</span></th>
								<td class="left">
									<select name="emp_pay_id" id="emp_pay_id" style="width:90px;">
										<option value="">����</option>
										<option value='0'>����</option>
										<option value='1'>����</option>
										<option value='2'>����</option>
										<option value='3'>¡��</option>
										<option value='5'>����</option>
									</select>
								</td>
								<th>�޿� ID<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="dz_id" id="dz_id" style="width:90px;" maxlength="7" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
							</tr>
						</tbody>
					</table>
				</div>
				<br/>
				<div align="center">
        			<span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();"/></span>
					<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>
<%

%>