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
Dim u_type, view_condi
Dim curr_date, curr_hh, curr_mm, t_date
Dim code_last, emp_reg_user, emp_mod_user, emp_name, emp_ename, emp_type, emp_sex
Dim emp_person1, emp_person2, emp_image, emp_first_date, emp_in_date, emp_gunsok_date
Dim emp_yuncha_date, emp_end_gisan, emp_end_date, emp_bonbu, emp_saupbu, emp_team
Dim emp_org_code, emp_org_name, emp_org_baldate, emp_stay_code, emp_stay_name
Dim emp_reside_place, emp_reside_company, emp_grade, emp_grade_date, emp_job, emp_position
Dim emp_jikgun, emp_jikmu, emp_birthday, emp_birthday_id, emp_family_zip, emp_family_sido
Dim emp_family_gugun, emp_family_dong, emp_family_addr, emp_zipcode, emp_sido, emp_gugun
Dim emp_dong, emp_addr, emp_tel_ddd, emp_tel_no1, emp_tel_no2, emp_hp_ddd, emp_hp_no1, emp_hp_no2
Dim emp_email, emp_military_id, emp_military_date1, emp_military_date2, emp_military_grade
Dim emp_military_comm, emp_hobby, emp_faith, emp_last_edu, emp_marry_date, emp_disabled
Dim emp_disab_grade, emp_sawo_date, emp_emergency_tel, emp_extension_no, emp_nation_code
Dim att_file, emp_sawo_id
Dim emp_reg_date, title_line, rsEmp, rsMemb, cost_center, cost_group, emp_pay_id, emp_mod_date
Dim photo_image, dz_id, rsMax, max_seq
Dim mg_org, emp_org_level, rs_etc

' �Է¹޾� ����Ÿ�� ��Ƶ� �ʵ��̸��� ���ǿ� �⺻���� null�� ����Ѱ�
u_type = f_Request("u_type")
emp_no = f_Request("emp_no")
view_condi = f_Request("view_condi")

code_last = ""
emp_reg_user = ""
emp_mod_user = ""

emp_name = ""
emp_ename = ""
emp_type = ""
emp_sex = ""
emp_person1 = ""
emp_person2 = ""
emp_image = ""
emp_first_date = ""
emp_in_date = ""
emp_gunsok_date = ""
emp_yuncha_date = ""
emp_end_gisan = ""
emp_end_date = ""
emp_company = ""
emp_bonbu = ""
emp_saupbu = ""
emp_team = ""
emp_org_code = ""
emp_org_name = ""
emp_org_baldate = ""
emp_stay_code = ""
emp_stay_name = ""
emp_reside_place = ""
emp_reside_company = ""
emp_grade = ""
emp_grade_date = ""
emp_job = ""
emp_position = ""
emp_jikgun = ""
emp_jikmu = ""
emp_birthday = ""
emp_birthday_id = ""
emp_family_zip = ""
emp_family_sido = ""
emp_family_gugun = ""
emp_family_dong = ""
emp_family_addr = ""
emp_zipcode = ""
emp_sido = ""
emp_gugun = ""
emp_dong = ""
emp_addr = ""
emp_tel_ddd = ""
emp_tel_no1 = ""
emp_tel_no2 = ""
emp_hp_ddd = ""
emp_hp_no1 = ""
emp_hp_no2 = ""
emp_email = ""
emp_military_id = ""
emp_military_date1 = ""
emp_military_date2 = ""
emp_military_grade = ""
emp_military_comm = ""
emp_hobby = ""
emp_faith = ""
emp_last_edu = ""
emp_marry_date = ""
emp_disabled = ""
emp_disab_grade = ""
emp_sawo_date = ""
emp_emergency_tel = ""
emp_extension_no = ""
emp_nation_code = ""
att_file = ""
emp_reg_user = user_name
emp_sawo_id = "N"

'first_date = curr_date
'request_hh = curr_hh
'request_mm = curr_mm
'emp_reg_date = Now()

title_line = "�λ�⺻���� ���"

If u_type = "U" Then
	objBuilder.Append "SELECT emtt.emp_name, emtt.emp_ename, emtt.emp_type, emtt.emp_sex, emtt.emp_person1, emtt.emp_person2, "
	objBuilder.Append "	emtt.emp_image, emtt.emp_first_date, emtt.emp_in_date, emtt.emp_gunsok_date, emtt.emp_yuncha_date, "
	objBuilder.Append "	emtt.emp_end_gisan, emtt.emp_end_date, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
	objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_org_baldate, emtt.emp_stay_code, emtt.emp_stay_name, "
	objBuilder.Append "	emtt.emp_reside_place, emtt.emp_reside_company, emtt.emp_grade, emtt.emp_grade_date, emtt.emp_job, "
	objBuilder.Append "	emtt.emp_position, emtt.emp_jikgun, emtt.emp_jikmu, emtt.emp_birthday, emtt.emp_birthday_id, emtt.emp_family_zip, "
	objBuilder.Append "	emtt.emp_family_sido, emtt.emp_family_gugun, emtt.emp_family_dong, emtt.emp_family_addr, emtt.emp_zipcode, "
	objBuilder.Append "	emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, "
	objBuilder.Append "	emtt.emp_hp_ddd, emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_email, emtt.emp_military_id, emtt.emp_military_date1, "
	objBuilder.Append "	emtt.emp_military_date2, emtt.emp_military_grade, emtt.emp_military_comm, emtt.emp_hobby, emtt.emp_faith, "
	objBuilder.Append "	emtt.emp_last_edu, emtt.emp_marry_date, emtt.emp_disabled, emtt.emp_disab_grade, emtt.emp_sawo_id, emtt.emp_sawo_date, "
	objBuilder.Append "	emtt.emp_emergency_tel, emtt.emp_nation_code, emtt.emp_pay_id, emtt.emp_extension_no, emtt.cost_center, "
	objBuilder.Append "	emtt.cost_group, emtt.emp_reg_date, emtt.emp_reg_user, emtt.emp_mod_date, emtt.emp_mod_user, emtt.emp_image, "
	objBuilder.Append "	dpit.dz_id "
	objBuilder.Append "FROM emp_master AS emtt "
	objBuilder.Append "LEFT OUTER JOIN dz_pay_info AS dpit ON emtt.emp_no = dpit.emp_no "
	objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"';"

	Set rsEmp = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	emp_name = rsEmp("emp_name")
    emp_ename = rsEmp("emp_ename")
    emp_type = rsEmp("emp_type")
    emp_sex = rsEmp("emp_sex")
    emp_person1 = rsEmp("emp_person1")
    emp_person2 = rsEmp("emp_person2")
    emp_image = rsEmp("emp_image")
	att_file = rsEmp("emp_image")
    emp_first_date = rsEmp("emp_first_date")
    emp_in_date = rsEmp("emp_in_date")
    emp_gunsok_date = rsEmp("emp_gunsok_date")
    emp_yuncha_date = rsEmp("emp_yuncha_date")
    emp_end_gisan = rsEmp("emp_end_gisan")
    emp_end_date = rsEmp("emp_end_date")
    emp_company = rsEmp("emp_company")
    emp_bonbu = rsEmp("emp_bonbu")
    emp_saupbu = rsEmp("emp_saupbu")
    emp_team = rsEmp("emp_team")
    emp_org_code = rsEmp("emp_org_code")
    emp_org_name = rsEmp("emp_org_name")
    emp_org_baldate = rsEmp("emp_org_baldate")
    emp_stay_code = rsEmp("emp_stay_code")
	emp_stay_name = rsEmp("emp_stay_name")
    emp_reside_place = rsEmp("emp_reside_place")
	emp_reside_company = rsEmp("emp_reside_company")
    emp_grade = rsEmp("emp_grade")
    emp_grade_date = rsEmp("emp_grade_date")
    emp_job = rsEmp("emp_job")
    emp_position = rsEmp("emp_position")
    emp_jikgun = rsEmp("emp_jikgun")
    emp_jikmu = rsEmp("emp_jikmu")
    emp_birthday = rsEmp("emp_birthday")
    emp_birthday_id = rsEmp("emp_birthday_id")
    emp_family_zip = rsEmp("emp_family_zip")
    emp_family_sido = rsEmp("emp_family_sido")
    emp_family_gugun = rsEmp("emp_family_gugun")
    emp_family_dong = rsEmp("emp_family_dong")
    emp_family_addr = rsEmp("emp_family_addr")
    emp_zipcode = rsEmp("emp_zipcode")
    emp_sido = rsEmp("emp_sido")
    emp_gugun = rsEmp("emp_gugun")
    emp_dong = rsEmp("emp_dong")
    emp_addr = rsEmp("emp_addr")
    emp_tel_ddd = rsEmp("emp_tel_ddd")
    emp_tel_no1 = rsEmp("emp_tel_no1")
    emp_tel_no2 = rsEmp("emp_tel_no2")
    emp_hp_ddd = rsEmp("emp_hp_ddd")
    emp_hp_no1 = rsEmp("emp_hp_no1")
    emp_hp_no2 = rsEmp("emp_hp_no2")
    emp_email = rsEmp("emp_email")
    emp_military_id = rsEmp("emp_military_id")
    emp_military_date1 = rsEmp("emp_military_date1")
    emp_military_date2 = rsEmp("emp_military_date2")
    emp_military_grade = rsEmp("emp_military_grade")
    emp_military_comm = rsEmp("emp_military_comm")
    emp_hobby = rsEmp("emp_hobby")
    emp_faith = rsEmp("emp_faith")
    emp_last_edu = rsEmp("emp_last_edu")
    emp_marry_date = rsEmp("emp_marry_date")
    emp_disabled = rsEmp("emp_disabled")
    emp_disab_grade = rsEmp("emp_disab_grade")
    emp_sawo_id = rsEmp("emp_sawo_id")

	If rsEmp("emp_sawo_id") = "" Or isNull(emp_sawo_id) Then
	   emp_sawo_id = "N"
	End If

    emp_sawo_date = rsEmp("emp_sawo_date")
    emp_emergency_tel = rsEmp("emp_emergency_tel")
    emp_nation_code = rsEmp("emp_nation_code")
	emp_extension_no = rsEmp("emp_extension_no")
	cost_center = rsEmp("cost_center")
	cost_group = rsEmp("cost_group")
	emp_pay_id = rsEmp("emp_pay_id")
'   end_date = mid(cstr(now()),1,10)
    emp_reg_date = rsEmp("emp_reg_date")
    emp_reg_user = rsEmp("emp_reg_user")
	emp_mod_date = rsEmp("emp_mod_date")
    emp_mod_user = rsEmp("emp_mod_user")

	photo_image = "/emp_photo/" & rsEmp("emp_image")
    att_file = rsEmp("emp_image")

	dz_id = rsEmp("dz_id")

	If rsEmp("emp_military_date1") = "1900-01-01" Then
		emp_military_date1 = ""
		emp_military_date2 = ""
    End If

	If rsEmp("emp_birthday") = "1900-01-01" Then
		emp_birthday = ""
    End If

    If rsEmp("emp_marry_date") = "1900-01-01" Then
		emp_marry_date = ""
    End If

	If rsEmp("emp_grade_date") = "1900-01-01" Then
		emp_grade_date = ""
    End If

	If rsEmp("emp_end_date") = "1900-01-01" Then
		emp_end_date = ""
    End If

	If rsEmp("emp_org_baldate") = "1900-01-01" Then
		emp_org_baldate = ""
    End If

	If rsEmp("emp_sawo_date") = "1900-01-01" Then
		emp_sawo_date = ""
    End If

	rsEmp.Close() : Set rsEmp = Nothing

	objBuilder.Append "SELECT mg_group FROM memb WHERE user_id= '"&emp_no&"';"

	Set rsMemb = DBConn.execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not rsMemb.eof Then
		mg_group = rsMemb("mg_group")
	Else
		mg_group = "1"
    End If
	rsMemb.Close() : Set rsMemb = Nothing

	title_line = "[ �λ�⺻���� ���� ]"
End If

objBuilder.Append "SELECT MAX(emp_no) AS max_seq FROM emp_master WHERE emp_no < '900000';"

Set rsMax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsMax("max_seq")) Then
	code_last = "000001"
Else
	max_seq = "000000" & CStr((Int(rsMax("max_seq")) + 1))
	code_last = Right(max_seq, 6)
End If

rsMax.Close() : Set rsMax = Nothing

If u_type = "U" Then
   code_last = emp_no
End If

emp_no = code_last
mg_group = "1"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				$( "#datepicker" ).datepicker("setDate", "<%=emp_first_date%>" );
			});

			//�Ի���
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=emp_in_date%>" );
			});

			//���������
			$(function(){
				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=emp_end_gisan%>" );
			});

			//�ټӱ����
			$(function(){
				$( "#datepicker3" ).datepicker();
				$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker3" ).datepicker("setDate", "<%=emp_gunsok_date%>" );
			});

			//���������
			$(function(){
				$( "#datepicker4" ).datepicker();
				$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker4" ).datepicker("setDate", "<%=emp_yuncha_date%>" );
			});

			//�������
			$(function(){
				$( "#datepicker5" ).datepicker();
				$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker5" ).datepicker("setDate", "<%=emp_birthday%>" );
			});

			//����������
			$(function(){
				$( "#datepicker6" ).datepicker();
				$( "#datepicker6" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker6" ).datepicker("setDate", "<%=emp_sawo_date%>" );
			});

			//��ȥ�����
			$(function(){
				$( "#datepicker7" ).datepicker();
				$( "#datepicker7" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker7" ).datepicker("setDate", "<%=emp_marry_date%>" );
			});

			//��������������
			$(function(){
				$( "#datepicker8" ).datepicker();
				$( "#datepicker8" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker8" ).datepicker("setDate", "<%=emp_military_date1%>" );
			});

			//��������������
			$(function(){
				$( "#datepicker9" ).datepicker();
				$( "#datepicker9" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker9" ).datepicker("setDate", "<%=emp_military_date2%>" );
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
				if(!$('#emp_no').val()){
					alert('��� ��ȣ�� �Է����ּ���');
					$('#emp_no').focus();
					return false;
				}

				if(document.frm.emp_name.value == ""){
					alert('������ �Է��ϼ���');
					frm.emp_name.focus();
					return false;
				}

				if(document.frm.emp_ename.value == ""){
					alert('���������� �Է����ּ���');
					frm.emp_ename.focus();
					return false;
				}

				if(document.frm.emp_birthday.value == ""){
					alert('��������� �Է����ּ���');
					frm.emp_birthday.focus();
					return false;
				}

				if(document.frm.emp_org_code.value == ""){
					alert('�Ҽ��� �������ּ���');
					frm.emp_org_code.focus();
					return false;
				}

				if(document.frm.emp_type.value == ""){
					alert('���������� �������ּ���');
					frm.emp_type.focus();
					return false;
				}

				if(document.frm.emp_grade.value == ""){
					alert('������ �������ּ���');
					frm.emp_grade.focus();
					return false;
				}

				if(document.frm.emp_job.value ==""){
					alert('������ �������ּ���');
					frm.emp_job.focus();
					return false;
				}

				if(document.frm.emp_position.value == ""){
					alert('��å�� �������ּ���');
					frm.emp_position.focus();
					return false;
				}

				if(document.frm.emp_jikmu.value == ""){
					alert('������ �������ּ���');
					frm.emp_jikmu.focus();
					return false;
				}

				if(document.frm.emp_first_date.value == ""){
					alert('�����Ի����� �Է����ּ���');
					frm.emp_first_date.focus();
					return false;
				}

				if(document.frm.emp_in_date.value == ""){
					alert('�Ի����� �Է����ּ���');
					frm.emp_in_date.focus();
					return false;
				}

				if(document.frm.emp_end_gisan.value == ""){
					alert('����������� �Է����ּ���');
					frm.emp_end_gisan.focus();
					return false;
				}

				if(document.frm.emp_gunsok_date.value == ""){
					alert('�ټӱ������ �Է����ּ���');
					frm.emp_gunsok_date.focus();
					return false;
				}

				if(document.frm.emp_yuncha_date.value == ""){
					alert('����������� �Է����ּ���');
					frm.emp_yuncha_date.focus();
					return false;
				}

				if(document.frm.emp_person1.value == ""){
					alert('�ֹε�Ϲ�ȣ�� �Է����ּ���');
					frm.emp_person1.focus();
					return false;
				}

				if(document.frm.emp_person2.value == ""){
					alert('�ֹε�Ϲ�ȣ�� �Է����ּ���');
					frm.emp_person2.focus();
					return false;
				}

				if(document.frm.emp_hp_ddd.value == ""){
					alert('�޴�����ȣ�� �Է����ּ���');
					return false;
				}

				if(document.frm.emp_hp_no1.value == ""){
					alert('�޴�����ȣ�� �Է����ּ���');
					return false;
				}

				if(document.frm.emp_hp_no2.value == ""){
					alert('�޴�����ȣ�� �Է����ּ���');
					return false;
				}

				if(isEmpty($('#emp_sido').val())){
					alert('�ּ�(��)�� ��ȸ���ּ���');
					return false;
				}
				if(document.frm.cost_center.value == ""){
					alert('��뱸���� �������ּ���');
					frm.cost_center.focus();
					return false;
				}
				/*
				//�λ��� ��û���� ���� ���� �ּ� ó��[����ȣ_20220613]
				if(document.frm.emp_first_date.value > document.frm.emp_in_date.value){
					alert('�����Ի����� �Ի��Ϻ��� �ʽ��ϴ�');
					frm.emp_first_date.focus();
					return false;
				}

				if(document.frm.emp_in_date.value > document.frm.emp_end_gisan.value){
					alert('����������� �Ի��Ϻ��� �����ϴ�');
					frm.emp_end_gisan.focus();
					return false;
				}

				if(document.frm.emp_in_date.value > document.frm.emp_yuncha_date.value){
					alert('����������� �Ի��Ϻ��� �����ϴ�');
					frm.emp_yuncha_date.focus();
					return false;
				}
				*/

				if(document.frm.emp_military_id.value !== ""){
					if(document.frm.emp_military_date1.value =="") {
						alert('���� ���� �Ⱓ�� �Է����ּ���');
						frm.emp_military_date1.focus();
						return false;
					}
				}

				if(document.frm.cost_center.value == "����������"){
					if(document.frm.emp_reside_company.value == ""){
						alert('����óȸ�縦 �����ϼ���');
						frm.emp_reside_company.focus();
						return false;
					}
				}

				if(!$('#dz_id').val()){
					alert('�޿�ID�� �Է����ּ���.');
					$('#dz_id').focus();
					return false;
				}

				var result = confirm('����Ͻðڽ��ϱ�?');

				if(result == true){
					return true;
				}
				return false;
			}

			function file_browse(){
           		document.frm.att_file.click();
           		document.frm.text1.value=document.frm.att_file.value;
			}

			// opener���� ������ �߻��ϴ� ��� �Ʒ� �ּ��� �����ϰ�, ������� ������������ �Է��մϴ�. ("�˾�API ȣ�� �ҽ�"�� �����ϰ� ������Ѿ� �մϴ�.)
			//document.domain = "abc.go.kr";
			function jusoCallBack(roadFullAddr,roadAddrPart1,addrDetail,roadAddrPart2,engAddr,jibunAddr,zipNo,admCd,rnMgtSn,bdMgtSn,detBdNmList,bdNm,bdKdcd,siNm,sggNm,emdNm,liNm,rn,udrtYn,buldMnnm,buldSlno,mtYn,lnbrMnnm,lnbrSlno,emdNo,gubun){
				/*document.getElementById('roadFullAddr').value = roadFullAddr;
				document.getElementById('roadAddrPart1').value = roadAddrPart1;
				document.getElementById('addrDetail').value = addrDetail;
				document.getElementById('roadAddrPart2').value = roadAddrPart2;
				document.getElementById('engAddr').value = engAddr;
				document.getElementById('jibunAddr').value = jibunAddr;
				document.getElementById('zipNo').value = zipNo;
				document.getElementById('admCd').value = admCd;
				document.getElementById('rnMgtSn').value = rnMgtSn;
				document.getElementById('bdMgtSn').value = bdMgtSn;
				document.getElementById('detBdNmList').value = detBdNmList;
				*/
				/**2017�� 2�� �߰� ���� **/
				/*
				document.getElementById('bdNm').value = bdNm;
				document.getElementById('bdKdcd').value = bdKdcd;
				document.getElementById('siNm').value = siNm;
				document.getElementById('sggNm').value = sggNm;
				document.getElementById('emdNm').value = emdNm;
				document.getElementById('liNm').value = liNm;
				document.getElementById('rn').value = rn;
				document.getElementById('udrtYn').value = udrtYn;
				document.getElementById('buldMnnm').value = buldMnnm;
				document.getElementById('buldSlno').value = buldSlno;
				document.getElementById('mtYn').value = mtYn;
				document.getElementById('lnbrMnnm').value = lnbrMnnm;
				document.getElementById('lnbrSlno').value = lnbrSlno;
				*/
				/**2017�� 3�� �߰� ���� **/
				//document.getElementById('emdNo').value = emdNo;

				//console.log(gubun);

				if(gubun === 'juso'){
					$('#emp_sido').val(siNm);
					$('#emp_gugun').val(sggNm);
					$('#emp_dong').val(rn + ' ' + buldMnnm);
					$('#emp_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#emp_zipcode').val(zipNo);
				}else if(gubun === 'family'){
					$('#emp_family_sido').val(siNm);
					$('#emp_family_gugun').val(sggNm);
					$('#emp_family_dong').val(rn + ' ' + buldMnnm);
					$('#emp_family_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#emp_family_zip').val(zipNo);
				}
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
	<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_emp_add01_save.asp" method="post" name="frm" enctype="multipart/form-data">
					<input type="hidden" name="emp_reside_place" id="emp_reside_place" style="width:120px;" value="<%=emp_reside_place%>"/>
					<input type="hidden" name="emp_org_level" id="emp_org_level" style="width:120px;" value="<%=emp_org_level%>"/>
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
									<input type="text" name="emp_no" id="emp_no" style="width:80px;" value="<%=emp_no%>" class="no-input" readonly/>
								</td>
                                <th>����(�ѱ�)<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="emp_name" id="emp_name" size="9" value="<%=emp_name%>"/></td>
								<th>����(����)<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="emp_ename" id="emp_ename" style="width:160px" maxlength="20" value="<%=emp_ename%>"/>
								</td>
                                <th>�������<span style="color:red;">*</span></th>
                                <td colspan="2" class="left">
									<input type="text" name="emp_birthday" size="10" id="datepicker5" style="width:70px;" value="<%=emp_birthday%>"/>
									&nbsp;��&nbsp;
									<input type="radio" name="emp_birthday_id" value="��" <%If emp_birthday_id = "��" Then %>checked<%End If %>/>��
              						<input type="radio" name="emp_birthday_id" value="��" <%If emp_birthday_id = "��" Then %>checked<%End If %>/>��
                                </td>
                            </tr>
							<tr>
                                <th>�Ҽ�<span style="color:red;">*</span></th>
								<td colspan="3" class="left">
									<input type="text" name="emp_org_code" id="emp_org_code" style="width:40px;" value="<%=emp_org_code%>" readonly/>
									&nbsp;��&nbsp;
									<input type="text" name="emp_org_name" id="emp_org_name" style="width:120px;" value="<%=emp_org_name%>" readonly/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=org&mg_org=<%=mg_org%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
                                </td>
                                <th>����<span style="color:red;">*</span></th>
                                <td colspan="5" class="left">
									<input type="text" name="emp_company" id="emp_company" style="width:100px;" value="<%=emp_company%>" readonly/>
									<input type="text" name="emp_bonbu" id="emp_bonbu" style="width:120px;" value="<%=emp_bonbu%>" readonly/>
									<input type="text" name="emp_saupbu" id="emp_saupbu" style="width:120px;" value="<%=emp_saupbu%>" readonly/>
									<input type="text" name="emp_team" id="emp_team" style="width:120px;" value="<%=emp_team%>" readonly/>
                                </td>
                            </tr>
							<tr>
                                <th>��������<span style="color:red;">*</span></th>
                                <td class="left">
									<select name="emp_type" id="emp_type" value="<%=emp_type%>" style="width:90px;">
										<option value="" <%If emp_type = "" Then %>selected<%End If %>>����</option>
										<option value='����' <%If emp_type = "����" Then %>selected<%End If %>>����</option>
										<option value='����' <%If emp_type = "����" Then %>selected<%End If %>>����</option>
										<option value='�����' <%If emp_type = "�����" Then %>selected<%End If %>>�����</option>
									</select>
                                </td>
                               	<th>����<span style="color:red;">*</span></th>
								<td class="left">
								<%
								Dim rsGrade, rsJob, rsPosition, rsJikmu

								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '02' ORDER BY emp_etc_code ASC;"

								Set rsGrade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="emp_grade" id="emp_grade" style="width:90px;">
										<option value="" <%If emp_grade = "" Then %>selected<%End If %>>����</option>
                				<%
								Do Until rsGrade.EOF
			  					%>
                						<option value='<%=rsGrade("emp_etc_name")%>' <%If emp_grade = rsGrade("emp_etc_name") Then %>selected<%End If %>><%=rsGrade("emp_etc_name")%></option>
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
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '03' ORDER BY emp_etc_code ASC;"

								Set rsJob = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="emp_job" id="emp_job" style="width:90px">
										<option value="" <%If emp_job = "" Then %>selected<%End If %>>����</option>
                				<%
								Do Until rsJob.EOF
			  					%>
                						<option value='<%=rsJob("emp_etc_name")%>' <%If emp_job = rsJob("emp_etc_name") Then %>selected<%End If %>><%=rsJob("emp_etc_name")%></option>
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
									<select name="emp_position" id="emp_position" style="width:90px;">
										<option value="" <%If emp_position = "" Then %>selected<%End If %>>����</option>
                				<%
								Do Until rsPosition.EOF
			  					%>
                						<option value='<%=rsPosition("emp_etc_name")%>' <%If emp_position = rsPosition("emp_etc_name") Then %>selected<%End If %>><%=rsPosition("emp_etc_name")%></option>
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
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '05' ORDER BY emp_etc_code ASC;"

								Set rsJikmu = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="emp_jikmu" id="emp_jikmu" style="width:90px;">
										<option value="" <%If emp_jikmu = "" Then %>selected<%End If %>>����</option>
                				<%
								Do Until rsJikmu.EOF
			  					%>
                					<option value='<%=rsJikmu("emp_etc_name")%>' <%If emp_jikmu = rsJikmu("emp_etc_name") Then %>selected<%End If %>><%=rsJikmu("emp_etc_name")%></option>
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
									<input type="text" name="emp_first_date" size="10" id="datepicker" style="width:70px;" value="<%=emp_first_date%>"/>&nbsp;
                                </td>
                                <th>�Ի���<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="emp_in_date" size="10" id="datepicker1" style="width:70px;" value="<%=emp_in_date%>"/>&nbsp;
                                </td>
                                <th>���������<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="emp_end_gisan" size="10" id="datepicker2" style="width:70px;" value="<%=emp_end_gisan%>"/>
                                </td>
                                <th>�ټӱ����<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="emp_gunsok_date" size="10" id="datepicker3" style="width:70px;" value="<%=emp_gunsok_date%>"/>
                                </td>
                                <th>���������<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="emp_yuncha_date" size="10" id="datepicker4" style="width:70px;" value="<%=emp_yuncha_date%>"/>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="2">�ֹι�ȣ<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="emp_person1" id="emp_person1" style="width:40px;" maxlength="6" value="<%=emp_person1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									��
									<input type="text" name="emp_person2" id="emp_person2" style="width:50px;" maxlength="7" value="<%=emp_person2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									����
									<select name="emp_sex" id="emp_sex" value="<%=emp_sex%>" style="width:50px;">
										<option value="" <%If emp_sex = "" Then %>selected<%End If %>>����</option>
										<option value='��' <%If emp_sex = "��" Then %>selected<%End If %>>��</option>
										<option value='��' <%If emp_sex = "��" Then %>selected<%End If %>>��</option>
									</select>
                                </td>
                                <th>��ȭ��ȣ</th>
								<td colspan="3" class="left">
									<input type="text" name="emp_tel_ddd" id="emp_tel_ddd" size="3" maxlength="3" value="<%=emp_tel_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_tel_no1" id="emp_tel_no1" size="4" maxlength="4" value="<%=emp_tel_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_tel_no2" id="emp_tel_no2" size="4" maxlength="4" value="<%=emp_tel_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                </td>
                                <th>�޴���<span style="color:red;">*</span></th>
								<td colspan="3" class="left">
									<input type="text" name="emp_hp_ddd" id="emp_hp_ddd" size="3" maxlength="3" value="<%=emp_hp_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_hp_no1" id="emp_hp_no1" size="4" maxlength="4" value="<%=emp_hp_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_hp_no2" id="emp_hp_no2" size="4" maxlength="4" value="<%=emp_hp_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="2">��󿬶�</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_emergency_tel" id="emp_emergency_tel" style="width:100px;" value="<%=emp_emergency_tel%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
								<th>�����з�</th>
                                <td colspan="2" class="left">
                                <select name="emp_last_edu" id="emp_last_edu" value="<%=emp_last_edu%>" style="width:100px;">
			            	        <option value="" <%If emp_last_edu = "" Then %>selected<%End If %>>����</option>
				                    <option value='����б�' <%If emp_last_edu = "����б�" Then %>selected<%End If %>>����б�</option>
                                    <option value='������' <%If emp_last_edu = "������" Then %>selected<%End If %>>������</option>
                                    <option value='���б�' <%If emp_last_edu = "���б�" Then %>selected<%End If %>>���б�</option>
                                    <option value='���п�����' <%If emp_last_edu = "���п�����" Then %>selected<%End If %>>���п�����</option>
                                    <option value='���п�' <%If emp_last_edu = "���п�" Then %>selected<%End If %>>���п�</option>
                                </select>
                                </td>
								 <th>�̸��� �ּ�<span style="color:red;">*</span></th>
								<td colspan="4" class="left">
									<input type="text" name="emp_email" id="emp_email" size="12" value="<%=emp_email%>"/>
									@k-one.co.kr
                                </td>
                            </tr>
                            <tr>
								<th colspan="2">�ּ�(��)<span style="color:red;">*</span></th>
								<td colspan="10" class="left">
									<input type="text" name="emp_zipcode" id="emp_zipcode" style="width:50px;" value="<%=emp_zipcode%>"/>
									-
									<input type="text" name="emp_sido" id="emp_sido" style="width:100px;" readonly="true" value="<%=emp_sido%>"/>
									<input type="text" name="emp_gugun" id="emp_gugun" style="width:150px;" readonly="true" value="<%=emp_gugun%>"/>
									<input type="text" name="emp_dong" id="emp_dong" style="width:150px;" readonly="true" value="<%=emp_dong%>"/>
									<input type="text" name="emp_addr" id="emp_addr" style="width:200px;" value="<%=emp_addr%>" />

									<a href="#" class="btnType03" onClick="pop_Window('/insa/jusoPopup.asp?gubun=<%="juso"%>','family_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                                </td>
                            </tr>
                         	<tr>
								<th colspan="2" class="first">�������Կ���</th>
                                <td class="left">
									<input type="radio" name="emp_sawo_id" value="Y" <%If emp_sawo_id = "Y" Then %>checked<%End If %>/>����
              						<input type="radio" name="emp_sawo_id" value="N" <%If emp_sawo_id = "N" Then %>checked<%End If %>/>����
                                </td>
                                <th>����������</th>
                                <td class="left">
									<input type="text" name="emp_sawo_date" size="10" id="datepicker6" style="width:70px;" value="<%=emp_sawo_date%>"/>
                                </td>
								<th>��ȥ�����</th>
                                <td class="left">
									<input type="text" name="emp_marry_date" size="10" id="datepicker7" style="width:70px;" value="<%=emp_marry_date%>"/>
								</td>
								<th>���</th>
                                <td class="left">
									<input type="text" name="emp_hobby" id="emp_hobby" style="width:80px;" value="<%=emp_hobby%>"/>
								</td>
                                <th>���/���</th>
								<td colspan="2" class="left">
								<%
								Dim rsDisab, rsMilitaryId, rsMilitaryGrade

								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '22' ORDER BY emp_etc_code ASC;"

								Set rsDisab = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="emp_disabled" id="emp_disabled" style="width:100px;">
										<option value="" <%If emp_disabled = "" Then %>selected<%End If %>>����</option>
                				<%
								Do Until rsDisab.EOF
			  					%>
                					<option value='<%=rsDisab("emp_etc_name")%>' <%If emp_disabled = rsDisab("emp_etc_name") Then %>selected<%End If  %>><%=rsDisab("emp_etc_name")%></option>
                				<%
									rsDisab.MoveNext()
								Loop
								rsDisab.Close() : Set rsDisab = Nothing
								%>
            						</select>
									 -
									<select name="emp_disab_grade" id="emp_disab_grade" value="<%=emp_disab_grade%>" style="width:50px;">
										<option value="" <%If emp_disab_grade = "" Then %>selected<%End If %>>����</option>
										<option value='1��' <%If emp_disab_grade = "1��" Then %>selected<%End If %>>1��</option>
										<option value='2��' <%If emp_disab_grade = "2��" Then %>selected<%End If %>>2��</option>
										<option value='3��' <%If emp_disab_grade = "3��" Then %>selected<%End If %>>3��</option>
										<option value='4��' <%If emp_disab_grade = "4��" Then %>selected<%End If %>>4��</option>
										<option value='5��' <%If emp_disab_grade = "5��" Then %>selected<%End If %>>5��</option>
										<option value='6��' <%If emp_disab_grade = "6��" Then %>selected<%End If %>>6��</option>
										<option value='����' <%If emp_disab_grade = "����" Then %>selected<%End If %>>����</option>
										<option value='����' <%If emp_disab_grade = "����" Then %>selected<%End If %>>����</option>
									</select>
                                </td>
                 			</tr>
                            <tr>
                                <th colspan="2" >��������</th>
                                <td class="left">
								<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '06' ORDER BY emp_etc_code ASC;"

								Set rsMilitaryId = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="emp_military_id" id="emp_military_id" style="width:90px">
										<option value="" <% if emp_military_id = "" Then %>selected<%End If %>>����</option>
                				<%
								Do Until rsMilitaryId.EOF
			  					%>
                						<option value='<%=rsMilitaryId("emp_etc_name")%>' <%If emp_military_id = rsMilitaryId("emp_etc_name") Then %>selected<%End If  %>><%=rsMilitaryId("emp_etc_name")%></option>
                				<%
									rsMilitaryId.MoveNext()
								Loop
								rsMilitaryId.Close() : Set rsMilitaryId = Nothing
								%>
                					</select>
                                </td>
                                <th>�������</th>
                                <td class="left">
								<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '07' ORDER BY emp_etc_code ASC;"

								Set rsMilitaryGrade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="emp_military_grade" id="emp_military_grade" style="width:90px">
										<option value="" <% if emp_military_grade = "" then %>selected<% end if %>>����</option>
                				<%
								do until rsMilitaryGrade.eof
			  					%>
                						<option value='<%=rsMilitaryGrade("emp_etc_name")%>' <%If emp_military_grade = rsMilitaryGrade("emp_etc_name") Then %>selected<%End If %>><%=rsMilitaryGrade("emp_etc_name")%></option>
                				<%
									rsMilitaryGrade.MoveNext()
								Loop
								rsMilitaryGrade.Close() : Set rsMilitaryGrade = Nothing
								%>
                					</select>
                                </td>
                                <th>���� �����Ⱓ</th>
                                <td colspan="2" class="left">
									<input type="text" name="emp_military_date1" size="10" id="datepicker8" style="width:70px;" value="<%=emp_military_date1%>"/>
									��
									<input type="text" name="emp_military_date2" size="10" id="datepicker9" style="width:70px;" value="<%=emp_military_date2%>"/>
                                </td>
                                <th>��������</th>
								<td class="left">
									<input type="text" name="emp_military_comm" id="emp_military_comm" style="width:80px;" value="<%=emp_military_comm%>"/>
								</td>
                                <th>����</th>
                                <td class="left">
									<input type="text" name="emp_faith" id="emp_faith" style="width:90px;" value="<%=emp_faith%>"/>
								</td>
							</tr>
                            <tr>
                        		<th colspan="2" class="first">�Ǳٹ���/�ּ�</th>
                                <td colspan="2" class="left">
									<input type="text" name="emp_stay_name" id="emp_stay_name" size="15"  value="<%=emp_stay_name%>"/>

									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_stay_select.asp?gubun=<%="stay"%>&reside_code=<%=emp_stay_code%>','stayselect','scrollbars=yes,width=1000,height=400')">����</a>
								</td>
                                <td colspan="6" class="left">
								<%
								Dim stay_sido, stay_gugun, stay_dong, stay_addr

								If emp_stay_code <> "" Then
								   'Sql = "select * from emp_stay where stay_code = '"&emp_stay_code&"'"
								   objBuilder.Append "SELECT stay_name, stay_sido, stay_gugun, stay_dong, stay_addr "
								   objBuilder.Append "FROM emp_stay "
								   objBuilder.Append "WHERE stay_code = '"&emp_stay_code&"' "

								   Set rs_stay = DBConn.Execute(objBuilder.ToString())
								   objBuilder.Clear()

							    	'do until rs_stay.eof
							    	If Not rs_stay.EOF Then

								       emp_stay_name = rs_stay("stay_name")
								       stay_sido = rs_stay("stay_sido")
								       stay_gugun = rs_stay("stay_gugun")
								       stay_dong = rs_stay("stay_dong")
								       stay_addr = rs_stay("stay_addr")
								   '	rs_stay.movenext()
								    'loop
								     End If
								     rs_stay.Close() : Set rs_stay = Nothing
								End If
								%>
									<input type="text" name="emp_stay_code" id="emp_stay_code" size="4" value="<%=emp_stay_code%>" readonly/>
									~
									<input type="text" name="stay_sido" id="stay_sido" style="width:90px;" value="<%=stay_sido%>" readonly/>
									<input type="text" name="stay_gugun" id="stay_gugun" style="width:90px;" value="<%=stay_gugun%>" readonly/>
									<input type="text" name="stay_dong" id="stay_dong" style="width:90px;" value="<%=stay_dong%>" readonly/>
									<input type="text" name="stay_addr" id="stay_addr" style="width:150px;" value="<%=stay_addr%>" readonly/>
								</td>
								<th>���׷�</th>
								<td class="left">
									<input type="text" name="cost_group" id="cost_group" style="width:90px" value="<%=cost_group%>" readonly/>
            					</td>
                            </tr>
                            <tr>
                        		<th colspan="2" class="first">������ȣ</th>
                                <td colspan="2" class="left">
									<input type="text" name="emp_extension_no" id="emp_extension_no" size="16 " value="<%=emp_extension_no%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                </td>
								<th>����ó ȸ��</th>
								<td colspan="2" class="left">
									<input type="text" name="emp_reside_company" id="emp_reside_company" readonly="true" style="width:90px;" value="<%=emp_reside_company%>"/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_trade_search.asp?gubun=5','tradesearch','scrollbars=yes,width=600,height=400')">ã��</a>
            					</td>
                                <th>��뱸��</th>
                                <td class="left">
								<%
								Dim rsCostCenter

								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC;"

								Set rsCostCenter = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="cost_center" id="cost_center" style="width:90px">
										<option value="" <% if cost_center = "" then %>selected<% end if %>>����</option>
                				<%
								Do Until rsCostCenter.EOF
			  					%>
                						<option value='<%=rsCostCenter("emp_etc_name")%>' <%If cost_center = rsCostCenter("emp_etc_name") then %>selected<% end if %>><%=rsCostCenter("emp_etc_name")%></option>
                				<%
									rsCostCenter.MoveNext()
								Loop
								rsCostCenter.Close() : Set rsCostCenter = Nothing
								DBConn.Close() : Set DBConn = Nothing
								%>
                					</select>
                                </td>
                                <th>�����׷쿩��</th>
                                <td colspan="2" class="left">
									<input type="radio" name="mg_group" value="1" <%If mg_group = "1" Then %>checked<%End If %>/>�Ϲݱ׷�
              						<input type="radio" name="mg_group" value="2" <%If mg_group = "2" Then %>checked<%End If %>/>�����׷�
                                </td>
                            </tr>
                            <tr>
                        		<th colspan="2" class="first">�Է���</th>
                                <td colspan="2" class="left"><%=emp_reg_date%>&nbsp;(<%=emp_reg_user%>)</td>
                                <th>������</th>
                                <td colspan="2" class="left"><%=emp_mod_date%>&nbsp;(<%=emp_mod_user%>)</td>

                                <th>�޿����</th>
                                <td class="left">
									<select name="emp_pay_id" id="emp_pay_id" value="<%=emp_pay_id%>" style="width:90px;">
										<option value="" <%If emp_pay_id = "" Then %>selected<%End If %>>����</option>
										<option value='0' <%If emp_pay_id = "0" Then %>selected<%End If %>>����</option>
										<option value='1' <%If emp_pay_id = "1" Then %>selected<%End If %>>����</option>
										<option value='2' <%If emp_pay_id = "2" Then %>selected<%End If %>>����</option>
										<option value='3' <%If emp_pay_id = "3" Then %>selected<%End If %>>¡��</option>
										<option value='5' <%If emp_pay_id = "5" Then %>selected<%End If %>>����</option>
									</select>
                                </td>
								<th>�޿�ID<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="dz_id" id="dz_id" style="width:90px;" value="<%=dz_id%>" maxlength="7" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
                            </tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th scope="row">�������</th>
								<td class="left">
									<input type="file" name= "att_file"  size="70" accept="image/gif" /> * ÷�������� 1���� �����ϸ� �ִ�뷮�� 2MB
                                </td>
							</tr>
						</tbody>
                    </table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();"/></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>"/>
                <input type="hidden" name="view_condi" value="<%=view_condi%>"/>
                <input type="hidden" name="emp_end_date" value="<%=emp_end_date%>"/>
                <input type="hidden" name="emp_org_baldate" value="<%=emp_org_baldate%>"/>
                <input type="hidden" name="emp_grade_date" value="<%=emp_grade_date%>"/>
                <input type="hidden" name="v_att_file" value="<%=att_file%>"/>
                <input type="hidden" name="t_date" value="<%=t_date%>"/>
			</form>
		</div>
	</div>
	</body>
</html>