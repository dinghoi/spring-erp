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
'### Include Request & Params
'===================================================
Dim at_name, at_empno, at_position
'===================================================
'### Request & Params
'===================================================
Dim curr_date, t_date, u_type
Dim emp_name, be_pg

curr_date = mid(cstr(now()),1,10)

t_date = mid(cstr(now()),1,7) + "-" + "01"

u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
be_pg = request("be_pg")


Dim rs_etc, rs_emp, rs_org, rs_stay

Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")

Dim app_seq, app_id, app_date, app_id_type, app_to_company
Dim app_to_org, app_to_grade, app_to_job, app_to_enddate
Dim app_be_company, app_be_org, app_be_grade, app_be_job
Dim app_be_enddate, app_first_date, app_end_date, app_comment

app_seq = ""
app_id = ""
app_date = ""
app_id_type = ""
app_to_company = ""
app_to_org = ""
app_to_grade = ""
app_to_job = ""
'app_to_grade = ""
app_to_enddate = ""
app_be_company = ""
app_be_org = ""
app_be_grade = ""
app_be_job = ""
'app_be_grade = ""
app_be_enddate = ""
app_first_date = ""
app_end_date = ""
app_comment = ""

Dim emp_ename, emp_type, emp_sex, emp_person1, emp_person2
Dim emp_first_date, emp_in_date
Dim emp_gunsok_date, emp_yuncha_date, emp_end_gisan, emp_end_date
Dim emp_bonbu, emp_saupbu, emp_team, emp_org_code
Dim emp_org_name, emp_org_baldate, emp_stay_code, emp_stay_name
Dim emp_reside_place, emp_reside_company, emp_grade, emp_grade_date
Dim emp_job, emp_position, emp_jikgun, emp_jikmu, emp_birthday
Dim emp_birthday_id, emp_family_zip, emp_family_sido, emp_family_gugun
Dim emp_family_dong, emp_family_addr, emp_zipcode, emp_sido, emp_gugun
Dim emp_dong, emp_addr, emp_tel_ddd, emp_tel_no1, emp_tel_no2
Dim emp_hp_ddd,	emp_hp_no1, emp_hp_no2, emp_email, emp_military_id
Dim emp_military_date1,	emp_military_date2, emp_military_grade
Dim emp_military_comm, emp_hobby, emp_faith, emp_last_edu
Dim emp_marry_date, emp_disabled, emp_disab_grade, emp_sawo_id
Dim emp_sawo_date, emp_emergency_tel, emp_nation_code, cost_center
Dim cost_group, photo_image, title_line

if u_type = "U" then
	'Sql="select * from emp_master where emp_no = '"&emp_no&"'"
	objBuilder.Append "SELECT emp_name, emp_ename, emp_type, emp_sex, emp_person1, "
	objBuilder.Append "emp_person2, emp_image, emp_first_date, emp_in_date, "
	objBuilder.Append "emp_gunsok_date, emp_yuncha_date, emp_end_gisan, emp_end_date, "
	objBuilder.Append "emp_company, emp_bonbu, emp_saupbu, emp_team, emp_org_code, "
	objBuilder.Append "emp_org_name, emp_org_baldate, emp_stay_code, emp_stay_name, "
	objBuilder.Append "emp_reside_place, emp_reside_company, emp_grade, emp_grade_date, "
	objBuilder.Append "emp_job, emp_position, emp_jikgun, emp_jikmu, emp_birthday, "
	objBuilder.Append "emp_birthday_id, emp_family_zip, emp_family_sido, emp_family_gugun, "
	objBuilder.Append "emp_family_dong, emp_family_addr, emp_zipcode, emp_sido, emp_gugun, "
	objBuilder.Append "emp_dong, emp_addr, emp_tel_ddd, emp_tel_no1, emp_tel_no2, "
	objBuilder.Append "emp_hp_ddd,	emp_hp_no1, emp_hp_no2, emp_email, emp_military_id, "
	objBuilder.Append "emp_military_date1,	emp_military_date2, emp_military_grade, "
	objBuilder.Append "emp_military_comm, emp_hobby, emp_faith, emp_last_edu, "
	objBuilder.Append "emp_marry_date, emp_disabled, emp_disab_grade, emp_sawo_id, "
	objBuilder.Append "emp_sawo_date, emp_emergency_tel, emp_nation_code, cost_center, "
	objBuilder.Append "cost_group "
	objBuilder.Append "FROM emp_master "
	objBuilder.Append "WHERE emp_no = '"&emp_no&"'"

	Set rs_emp = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()




	emp_name = rs_emp("emp_name")
    emp_ename = rs_emp("emp_ename")
    emp_type = rs_emp("emp_type")
    emp_sex = rs_emp("emp_sex")
    emp_person1 = rs_emp("emp_person1")
    emp_person2 = rs_emp("emp_person2")
 '   emp_image = rs_emp("emp_image")
    emp_first_date = rs_emp("emp_first_date")
    emp_in_date = rs_emp("emp_in_date")
    emp_gunsok_date = rs_emp("emp_gunsok_date")
    emp_yuncha_date = rs_emp("emp_yuncha_date")
    emp_end_gisan = rs_emp("emp_end_gisan")
    emp_end_date = rs_emp("emp_end_date")
    emp_company = rs_emp("emp_company")
    emp_bonbu = rs_emp("emp_bonbu")
    emp_saupbu = rs_emp("emp_saupbu")
    emp_team = rs_emp("emp_team")
    emp_org_code = rs_emp("emp_org_code")
    emp_org_name = rs_emp("emp_org_name")
    'emp_org_baldate = rs_emp("emp_org_baldate")
    emp_stay_code = rs_emp("emp_stay_code")
	emp_stay_name = rs_emp("emp_stay_name")
    emp_reside_place = rs_emp("emp_reside_place")
	emp_reside_company = rs_emp("emp_reside_company")
    emp_grade = rs_emp("emp_grade")
    'emp_grade_date = rs_emp("emp_grade_date")
    emp_job = rs_emp("emp_job")
    emp_position = rs_emp("emp_position")
    emp_jikgun = rs_emp("emp_jikgun")
    emp_jikmu = rs_emp("emp_jikmu")
    'emp_birthday = rs_emp("emp_birthday")
    emp_birthday_id = rs_emp("emp_birthday_id")
    emp_family_zip = rs_emp("emp_family_zip")
    emp_family_sido = rs_emp("emp_family_sido")
    emp_family_gugun = rs_emp("emp_family_gugun")
    emp_family_dong = rs_emp("emp_family_dong")
    emp_family_addr = rs_emp("emp_family_addr")
    emp_zipcode = rs_emp("emp_zipcode")
    emp_sido = rs_emp("emp_sido")
    emp_gugun = rs_emp("emp_gugun")
    emp_dong = rs_emp("emp_dong")
    emp_addr = rs_emp("emp_addr")
    emp_tel_ddd = rs_emp("emp_tel_ddd")
    emp_tel_no1 = rs_emp("emp_tel_no1")
    emp_tel_no2 = rs_emp("emp_tel_no2")
    emp_hp_ddd = rs_emp("emp_hp_ddd")
    emp_hp_no1 = rs_emp("emp_hp_no1")
    emp_hp_no2 = rs_emp("emp_hp_no2")
    emp_email = rs_emp("emp_email")
    emp_military_id = rs_emp("emp_military_id")
    'emp_military_date1 = rs_emp("emp_military_date1")
    emp_military_date2 = rs_emp("emp_military_date2")
    emp_military_grade = rs_emp("emp_military_grade")
    emp_military_comm = rs_emp("emp_military_comm")
    emp_hobby = rs_emp("emp_hobby")
    emp_faith = rs_emp("emp_faith")
    emp_last_edu = rs_emp("emp_last_edu")
    'emp_marry_date = rs_emp("emp_marry_date")
    emp_disabled = rs_emp("emp_disabled")
    emp_disab_grade = rs_emp("emp_disab_grade")
    emp_sawo_id = rs_emp("emp_sawo_id")
    'emp_sawo_date = rs_emp("emp_sawo_date")
    emp_emergency_tel = rs_emp("emp_emergency_tel")
    emp_nation_code = rs_emp("emp_nation_code")
	cost_center = rs_emp("cost_center")
	cost_group = rs_emp("cost_group")

	photo_image = "/emp_photo/" + rs_emp("emp_image")
	emp_email = emp_email + "@k-won.co.kr"

	if rs_emp("emp_birthday") = "1900-01-01" then
	   emp_birthday = ""
	   else
	   emp_birthday = rs_emp("emp_birthday")
	end if
	if rs_emp("emp_org_baldate") = "1900-01-01" then
	   emp_org_baldate = ""
	   else
	   emp_org_baldate = rs_emp("emp_org_baldate")
	end if
	if rs_emp("emp_grade_date") = "1900-01-01" then
	   emp_grade_date = ""
	   else
	   emp_grade_date = rs_emp("emp_grade_date")
	end if
	if rs_emp("emp_sawo_date") = "1900-01-01" then
	   emp_sawo_date = ""
	   else
	   emp_sawo_date = rs_emp("emp_sawo_date")
	end if
	if emp_sawo_id = "" or isNull(emp_sawo_id) then
	   emp_sawo_id = "N"
	end if
	if rs_emp("emp_military_date1") = "1900-01-01" then
       emp_military_date1 = ""
       emp_military_date2 = ""
       else
       emp_military_date1 = rs_emp("emp_military_date1")
       emp_military_date2 = rs_emp("emp_military_date2")
    end if
    if rs_emp("emp_marry_date") = "1900-01-01" then
       emp_marry_date = ""
       else
   	   emp_marry_date = rs_emp("emp_marry_date")
    end if
end if

title_line = " �λ� �߷� ��� "
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
				return "2 1";
			}

			function goAction(){
			   window.close() ;
			}

			function goBefore(){
			   history.back() ;
			}

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%'=app_date%>" );
			});

			$(function() {
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%'=app_be_enddate%>" );
			});

			$(function(){
				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%'=app_hustart_date%>" );
			});

			$(function(){
				$( "#datepicker3" ).datepicker();
				$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker3" ).datepicker("setDate", "<%'=app_hufinish_date%>" );
			});
			$(function(){
				$( "#datepicker4" ).datepicker();
				$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker4" ).datepicker("setDate", "<%'=app_distart_date%>" );
			});

			$(function(){
				$( "#datepicker5" ).datepicker();
				$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker5" ).datepicker("setDate", "<%'=app_difinish_date%>" );
			});

			function frmcheck(){
				if (chkfrm() && formcheck(document.frm)) {
					document.frm.submit ();
				}
			}

			function chkfrm(){
//				var a_date = document.frm.app_date.options[document.frm.app_date.selectedIndex].value;
//				var c_date = document.frm.curr_date.options[document.frm.curr_date.selectedIndex].value;

//				alert(a_date);
//				alert(c_date);

				if(document.frm.app_date.value ==""){
					alert('�߷����� �Է��ϼ���');
					frm.app_date.focus();
					return false;}

//				if(document.frm.app_id.value !="�����߷�")
//				    if(document.frm.app_date.value < document.frm.t_date.value) {
//					    alert('�߷����� �����Դϴ�. Ȯ���Ͻʽÿ�!');
//					    frm.app_date.focus();
//					    return false;}

				if(document.frm.app_id.value =="") {
					alert('�߷ɱ����� �����ϼ���');
					frm.app_id.focus();
					return false;}

				if(document.frm.app_id.value =="�̵��߷�")
					if(document.frm.app_be_orgcode.value =="") {
						alert('�߷ɼҼ��� �Է��ϼ���');
						frm.app_be_orgcode.focus();
						return false;}
				if(document.frm.app_id.value =="�̵��߷�")
					if(document.frm.app_mv_comment.value =="") {
						alert('�߷ɻ����� �Է��ϼ���');
						frm.app_mv_comment.focus();
						return false;}

				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_date.value =="") {
						alert('�������� �Է��ϼ���');
						frm.app_date.focus();
						return false;}
				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_end_type.value =="") {
						alert('���������� �����ϼ���');
						frm.app_end_type.focus();
						return false;}

				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_be_grade.value =="") {
						alert('���������� �����ϼ���');
						frm.app_be_grade.focus();
						return false;}
				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_gr_type.value =="") {
						alert('���������� �����ϼ���');
						frm.app_gr_type.focus();
						return false;}

				if(document.frm.app_id.value =="��å����")
					if(document.frm.app_be_position.value =="") {
						alert('������å�� �����ϼ���');
						frm.app_be_position.focus();
						return false;}
				if(document.frm.app_id.value =="��å����")
					if(document.frm.app_bm_orgcode.value =="") {
						alert('���ӼҼ��� �Է��ϼ���');
						frm.app_bm_orgcode.focus();
						return false;}

                if(document.frm.app_id.value =="��å����")
					if(document.frm.app_hm_type.value =="") {
						alert('���������� �����ϼ���');
						frm.app_hm_type.focus();
						return false;}

				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_hu_type.value =="") {
						alert('���������� �����ϼ���');
						frm.app_hu_type.focus();
						return false;}
				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_hustart_date.value =="") {
						alert('������������ �Է��ϼ���');
						frm.app_hustart_date.focus();
						return false;}
				if(document.frm.app_id.value =="�����߷�")
					if(document.frm.app_hufinish_date.value =="") {
						alert('������������ �Է��ϼ���');
						frm.app_hufinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="�����߷�")
				    if(document.frm.app_hustart_date.value > document.frm.app_hufinish_date.value) {
						alert('������������ ���������Ϻ��� �ʽ��ϴ�');
						frm.app_hufinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="�����߷�")
				    if(document.frm.app_date.value > document.frm.app_hustart_date.value) {
						alert('�߷��Ϻ��� ������������ �����ϴ�');
						frm.app_hustart_date.focus();
						return false;}

				if(document.frm.app_id.value =="¡��߷�")
					if(document.frm.app_di_type.value =="") {
						alert('¡�������� �����ϼ���');
						frm.app_di_type.focus();
						return false;}
				if(document.frm.app_id.value =="¡��߷�")
					if(document.frm.app_distart_date.value =="") {
						alert('¡��������� �Է��ϼ���');
						frm.app_distart_date.focus();
						return false;}
				if(document.frm.app_id.value =="¡��߷�")
					if(document.frm.app_difinish_date.value =="") {
						alert('¡�踸������ �Է��ϼ���');
						frm.app_difinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="¡��߷�")
				    if(document.frm.app_distart_date.value > document.frm.app_difinish_date.value) {
						alert('¡��������� ¡�踸���Ϻ��� �ʽ��ϴ�');
						frm.app_hufinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="¡��߷�")
				    if(document.frm.app_date.value > document.frm.app_distart_date.value) {
						alert('�߷��Ϻ��� ¡��������� �����ϴ�');
						frm.app_distart_date.focus();
						return false;}

				if(document.frm.app_id.value =="����߷�")
					if(document.frm.app_rw_type.value =="") {
						alert('���������� �����ϼ���');
						frm.app_rw_type.focus();
						return false;}
				if(document.frm.app_id.value =="����߷�")
					if(document.frm.app_rw_comment.value =="") {
						alert('���󳻿��� �Է��ϼ���');
						frm.app_rw_comment.focus();
						return false;}

				{
				a=confirm('�߷�ó���� �Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function menu1(){
				var c = document.frm.app_id.options[document.frm.app_id.selectedIndex].value;

				 {
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '�̵��߷�')
				{
					document.getElementById('mv_menu1').style.display = '';
					document.getElementById('mv_menu2').style.display = '';
					document.getElementById('mv_menu3').style.display = '';
					document.getElementById('mv_menu4').style.display = '';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '�����߷�')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = '';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '�����߷�')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = '';
					document.getElementById('gr_menu2').style.display = '';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '��å����')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = '';
					document.getElementById('bm_menu2').style.display = '';
					document.getElementById('bm_menu3').style.display = '';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '��å����')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = '';
					document.getElementById('hm_menu2').style.display = '';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '�����߷�')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = '';
					document.getElementById('hu_menu2').style.display = '';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '¡��߷�')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = '';
					document.getElementById('di_menu2').style.display = '';
					document.getElementById('rw_menu1').style.display = 'none';
				}
				if (c == '����߷�')
				{
					document.getElementById('mv_menu1').style.display = 'none';
					document.getElementById('mv_menu2').style.display = 'none';
					document.getElementById('mv_menu3').style.display = 'none';
					document.getElementById('mv_menu4').style.display = 'none';
					document.getElementById('end_menu1').style.display = 'none';
					document.getElementById('gr_menu1').style.display = 'none';
					document.getElementById('gr_menu2').style.display = 'none';
					document.getElementById('bm_menu1').style.display = 'none';
					document.getElementById('bm_menu2').style.display = 'none';
					document.getElementById('bm_menu3').style.display = 'none';
					document.getElementById('hm_menu1').style.display = 'none';
					document.getElementById('hm_menu2').style.display = 'none';
					document.getElementById('hu_menu1').style.display = 'none';
					document.getElementById('hu_menu2').style.display = 'none';
					document.getElementById('di_menu1').style.display = 'none';
					document.getElementById('di_menu2').style.display = 'none';
					document.getElementById('rw_menu1').style.display = '';
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_appoint_add_save.asp" method="post" name="frm">
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
                                <img src="<%=photo_image%>" width=110 height=120 alt="">
                                </td>
								<th>���&nbsp;&nbsp;��ȣ</th>
                                <td class="left"><%=emp_no%>&nbsp;</td>
                                <th>����(�ѱ�)</th>
                                <td class="left"><%=emp_name%>&nbsp;</td>
								<th>����(����)</th>
								<td colspan="2" class="left"><%=emp_ename%>&nbsp;</td>
                                <th>�������</th>
                                <td colspan="2" class="left"><%=emp_birthday%>&nbsp;(<%=emp_birthday_id%>)&nbsp;</td>
                           </tr>
                           <tr>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th>��������</th>
                                <td class="left"><%=emp_type%>&nbsp;</td>
                               	<th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_grade%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_job%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;å</th>
                                <td class="left"><%=emp_position%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_jikmu%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th>�����Ի���</th>
                                <td class="left"><%=emp_first_date%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</th>
                                <td class="left"><%=emp_in_date%>&nbsp;</td>
                                <th>���������</th>
                                <td class="left"><%=emp_end_gisan%>&nbsp;</td>
                                <th>�ټӱ����</th>
                                <td class="left"><%=emp_gunsok_date%>&nbsp;</td>
                                <th>���������</th>
                                <td class="left"><%=emp_yuncha_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">�ֹι�ȣ</th>
								<td colspan="2" class="left"><%=emp_person1%>&nbsp;��&nbsp;<%=emp_person2%>&nbsp;(<%=emp_sex%>)&nbsp;</td>
                                <th>��ȭ��ȣ</th>
								<td colspan="3" class="left"><%=emp_tel_ddd%>��<%=emp_tel_no1%>��<%=emp_tel_no2%>&nbsp;</td>
                                <th>�ڵ���</th>
								<td colspan="3" class="left"><%=emp_hp_ddd%>��<%=emp_hp_no1%>��<%=emp_hp_no2%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">���Ҽӹ߷���</th>
								<td class="left"><%=emp_org_baldate%>&nbsp;</td>
                                <th>�����޽�����</th>
								<td class="left"><%=emp_grade_date%>&nbsp;</td>
                                <th>e_�����ּ�</th>
								<td colspan="2" class="left"><%=emp_email%>&nbsp;</td>
                                <th>��������</th>
								<td colspan="3" class="left">
                                <input type="radio" name="emp_sawo_id" value="Y" <% if emp_sawo_id = "Y" then %>checked<% end if %>>����
              					<input name="emp_sawo_id" type="radio" value="N" <% if emp_sawo_id = "N" then %>checked<% end if %>>����
                                 &nbsp;&nbsp;��&nbsp;&nbsp;<%=emp_sawo_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="12" class="left" style="background:#FFC">�� �λ�&nbsp;&nbsp;&nbsp;�߷� ��</th>&nbsp;
                            </tr>
                            <tr>
                                <th colspan="2" class="first">�λ�߷�����</th>
                                <td colspan="3" class="left">
                                <input name="app_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;">&nbsp;</td>
                                <th class="first">�߷ɱ���</th>&nbsp;
                                <td colspan="6" class="left">
                            <%
								'Sql="select * from emp_etc_code where emp_etc_type = '10' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '10' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							%>
								<select name="app_id" id="select" style="width:150px" onChange="menu1()">
                                <option value="" <% if app_id = "" then %>selected<% end if %>>����</option>
                			<%
								do until rs_etc.eof
			  				%>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If app_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%>&nbsp;</option>
                			<%
									rs_etc.movenext()
								loop
								rs_etc.Close()
							%>
            					</select>
                                <input name="app_id_old" type="hidden" id="app_id_old" value="<%=app_id%>">
                                </td>
                            </tr>

							<!--�̵� �߷�-->
							<tr style="display:none;" id="mv_menu1">
								<th colspan="2" class="first" >���Ҽ�</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th class="first" >������</th>
                                <td colspan="6" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
							<%
							Dim app_be_orgcode, app_company, app_bonbu, app_saupbu, app_team
							Dim app_reside_place, emp_org_level, app_reside_company
							%>
                            <tr style="display:none;" id="mv_menu2">
								<th colspan="2" class="first" style="background:#FFC">�߷ɼҼ�</th>
								<td colspan="3" class="left">
								<input name="app_be_orgcode" type="text" id="app_be_orgcode" style="width:40px" readonly="true" value="<%=app_be_orgcode%>">
                                &nbsp;��&nbsp;
                                <input name="app_be_org" type="text" id="app_be_org" style="width:120px" readonly="true" value="<%=app_be_org%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="apporg"%>&view_condi=<%=emp_company%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
                                </td>
                                <th style="background:#FFC">�߷�����</th>
								<td colspan="6" class="left">
                                <input name="app_company" type="text" id="app_company" style="width:100px" readonly="true" value="<%=app_company%>">
              					<input name="app_bonbu" type="text" id="app_bonbu" style="width:120px" readonly="true" value="<%=app_bonbu%>">
              					<input name="app_saupbu" type="text" id="app_saupbu" style="width:120px" readonly="true" value="<%=app_saupbu%>">
              					<input name="app_team" type="text" id="app_team" style="width:120px" readonly="true" value="<%=app_team%>">
                                <input name="app_reside_place" type="hidden" id="app_reside_place" style="width:120px" readonly="true" value="<%=app_reside_place%>">&nbsp&nbsp&nbsp
                                <input name="app_reside_company" type="text" id="app_reside_company" style="width:120px" readonly="true" value="<%=app_reside_company%>">
                                <input name="app_org_level" type="hidden" id="emp_org_level" style="width:120px" readonly="true" value="<%=emp_org_level%>">
                                <input name="app_cost_group" type="hidden" id="app_cost_group" style="width:120px" readonly="true" value="<%=cost_group%>">
                                </td>
                            </tr>
                            <%
							Dim stay_name, stay_sido, stay_gugun, stay_dong, stay_addr

							stay_name = emp_stay_name

							if emp_stay_code <> "" then
								'Sql="select * from emp_stay where stay_code = '"&emp_stay_code&"'"
								objBuilder.Append "SELECT stay_name, stay_sido, stay_gugun, stay_dong, stay_addr "
								objBuilder.Append "FROM emp_stay "
								objBuilder.Append "WHERE stay_code = '"&emp_stay_code&"'"

								Rs_stay.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()

								'do until rs_stay.eof
								if not rs_stay.eof then

								   stay_name = rs_stay("stay_name")
								   stay_sido = rs_stay("stay_sido")
								   stay_gugun = rs_stay("stay_gugun")
								   stay_dong = rs_stay("stay_dong")
								   stay_addr = rs_stay("stay_addr")
							   '	rs_stay.movenext()
								'loop
								 end if
								 rs_stay.Close()
							end if
							%>

                            <tr style="display:none;" id="mv_menu3">
                                <th colspan="2" class="first" style="background:#FFC">�Ǳٹ���/�ּ�</th>
                                <td colspan="3" class="left">
                                <input name="emp_stay_code" type="text" id="emp_stay_code" style="width:40px" readonly="true" value="<%=emp_stay_code%>">
                                &nbsp;��&nbsp;
                                <input name="stay_name" type="text" id="stay_name" style="width:150px"  value="<%=stay_name%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_stay_select.asp?gubun=<%="stay"%>&reside_code=<%=emp_stay_code%>','stayselect','scrollbars=yes,width=1000,height=400')">����</a>
                                </td>
                                <th>�ּ���</th>
                                <td colspan="6" class="left">
                                <input name="stay_sido" type="text" id="stay_sido" style="width:100px" readonly="true" value="<%=stay_sido%>">
                                <input name="stay_gugun" type="text" id="stay_gugun" style="width:150px" readonly="true" value="<%=stay_gugun%>">
                                <input name="stay_dong" type="text" id="stay_dong" style="width:150px" readonly="true" value="<%=stay_dong%>">
                                <input name="stay_addr" type="text" id="stay_addr" style="width:200px" readonly="true" value="<%=stay_addr%>">
								</td>
                            </tr>
                            <tr style="display:none;" id="mv_menu4">
                                <th colspan="2" class="first" style="background:#FFC">����</th>
                                <td class="left">
                                <%
								'Sql="select * from emp_etc_code where emp_etc_type = '05' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '05' ORDER BY emp_etc_code ASC"

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
								<select name="emp_jikmu" id="emp_jikmu" style="width:90px">
                                <option value="" <% if emp_jikmu = "" then %>selected<% end if %>>����</option>
                			  <%
								do until rs_etc.eof
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_jikmu = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()
								loop
								rs_etc.Close()
							  %>
            					</select>
                                </td>
                                <th>��뱸��</th>
                                <td class="left">
                              <%
								'Sql="select * from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
								<select name="cost_center" id="cost_center" style="width:90px">
                                <option value="" <% if cost_center = "" then %>selected<% end if %>>����</option>
                			  <%
								do until rs_etc.eof
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If cost_center = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()
								loop
								rs_etc.Close()
							  %>
                			     </select>
                                </td>
								<th style="background:#FFC">�߷ɻ���</th>
								<td colspan="6" class="left">
								<input name="app_mv_comment" type="text" id="app_mv_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--���� �߷�-->
							<%
							Dim app_end_type
							%>
                            <tr style="display:none;" id="end_menu1">
                                <th colspan="2" style="background:#FFC">��������</th>
								<td colspan="2" class="left">
                                <select name="app_end_type" id="app_end_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_end_type = "" then %>selected<% end if %>>����</option>
				                    <option value='ȸ�����' <%If app_end_type = "ȸ�����" then %>selected<% end if %>>ȸ�����</option>
                                    <option value='������' <%If app_end_type = "������" then %>selected<% end if %>>������</option>
                                    <option value='���λ���' <%If app_end_type = "���λ���" then %>selected<% end if %>>���λ���</option>
                                    <option value='¡��' <%If app_end_type = "¡��" then %>selected<% end if %>>¡��</option>
                                    <option value='����' <%If app_end_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='����' <%If app_end_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='ġ��' <%If app_end_type = "ġ��" then %>selected<% end if %>>ġ��</option>
                                </select>
								<th style="background:#FFC">���� Comment.</th>
								<td colspan="7" class="left">
								<input name="app_end_comment" type="text" id="app_end_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--���� �߷�-->
							<%
							Dim app_gr_type
							%>
							<tr style="display:none;" id="gr_menu1">
								<th colspan="2" class="first" >������</th>
								<td class="left"><%=emp_grade%>&nbsp;</td>
                                <th>������</th>
								<td class="left"><%=emp_job%>&nbsp;</td>
                                <th>������ ������</th>
                                <td colspan="6" class="left"><%=emp_grade_date%>&nbsp;</td>
							</tr>
                            <tr style="display:none;" id="gr_menu2">
								<th colspan="2" class="first" style="background:#FFC">��������</th>
								<td class="left">
                            <%
								'Sql="select * from emp_etc_code where emp_etc_type = '02' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '02' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							%>
								<select name="app_be_grade" id="app_be_grade" style="width:90px">
                                <option value="" <% if app_be_grade = "" then %>selected<% end if %>>����</option>
                			<%
								do until rs_etc.eof
			  				%>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If app_be_grade = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%>&nbsp;</option>
                			<%
									rs_etc.movenext()
								loop
								rs_etc.Close()
							%>
            					</select>
                                </td>
                                <th style="background:#FFC">��������</th>
								<td class="left">
                                <select name="app_gr_type" id="app_gr_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_gr_type = "" then %>selected<% end if %>>����</option>
				                    <option value='�������' <%If app_gr_type = "�������" then %>selected<% end if %>>�������</option>
                                    <option value='Ư������' <%If app_gr_type = "Ư������" then %>selected<% end if %>>Ư������</option>
                                    <option value='���ǽ���' <%If app_gr_type = "���ǽ���" then %>selected<% end if %>>���ǽ���</option>
                                </select>
								<th style="background:#FFC">���� Comment.</th>
								<td colspan="6" class="left">
								<input name="app_gr_comment" type="text" id="app_gr_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--��å����-->
							<%
							Dim app_be_position, app_bm_orgcode, app_bm_org, app_bm_company, app_bm_bonbu
							Dim app_bm_saupbu, app_bm_team, app_bm_reside_place, app_bm_reside_company
							Dim app_bm_org_level
							%>
							<tr style="display:none;" id="bm_menu1">
                                <th colspan="2" class="first">��&nbsp;��å</th>
								<td class="left"><%=emp_position%>&nbsp;</td>
								<th >��&nbsp;�Ҽ�</th>
								<td colspan="2" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th >��&nbsp;����</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
                            <tr style="display:none;" id="bm_menu2">
								<th colspan="2" class="first" style="background:#FFC">������å</th>
								<td class="left">
                              <%
								'Sql="select * from emp_etc_code where emp_etc_type = '04' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '04' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
							  %>
								<select name="app_be_position" id="app_be_position" style="width:90px">
                                <option value="" <% if app_be_position = "" then %>selected<% end if %>>����</option>
                			  <%
								do until rs_etc.eof
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If app_be_position = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()
								loop
								rs_etc.Close()
							  %>
            					</select>
                                </td>
								<th style="background:#FFC">���ӼҼ�</th>
								<td colspan="2" class="left">
								<input name="app_bm_orgcode" type="text" id="app_bm_orgcode" style="width:30px" readonly="true" value="<%=app_bm_orgcode%>">
                                ��
                                <input name="app_bm_org" type="text" id="app_bm_org" style="width:80px" readonly="true" value="<%=app_bm_org%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="appbmorg"%>&view_condi=<%=emp_company%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
                                </td>
                                <th style="background:#FFC">��������</th>
								<td colspan="5" class="left">
                                <input name="app_bm_company" type="text" id="app_bm_company" style="width:100px" readonly="true" value="<%=app_bm_company%>">
              					<input name="app_bm_bonbu" type="text" id="app_bm_bonbu" style="width:120px" readonly="true" value="<%=app_bm_bonbu%>">
              					<input name="app_bm_saupbu" type="text" id="app_bm_saupbu" style="width:120px" readonly="true" value="<%=app_bm_saupbu%>">
              					<input name="app_bm_team" type="text" id="app_bm_team" style="width:120px" readonly="true" value="<%=app_bm_team%>">
                                <input name="app_bm_reside_place" type="hidden" id="app_bm_reside_place" style="width:120px" readonly="true" value="<%=app_bm_reside_place%>">
                                <input name="app_bm_reside_company" type="hidden" id="app_bm_reside_company" style="width:120px" readonly="true" value="<%=app_bm_reside_company%>">
                                <input name="app_bm_org_level" type="hidden" id="app_bm_org_level" style="width:120px" readonly="true" value="<%=app_bm_org_level%>">
                                </td>
							</tr>
                            <tr style="display:none;" id="bm_menu3">
								<th colspan="2" class="first" style="background:#FFC">���� Comment.</th>
								<td colspan="10" class="left">
								<input name="app_bm_comment" type="text" id="app_bm_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--��å����-->
							<%
							Dim app_hm_type, app_hu_type
							%>
							<tr style="display:none;" id="hm_menu1">
                                <th colspan="2" class="first">��&nbsp;��å</th>
								<td class="left"><%=emp_position%>&nbsp;</td>
								<th >��&nbsp;�Ҽ�</th>
								<td colspan="2" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th >��&nbsp;����</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
                            <tr style="display:none;" id="hm_menu2">
								<th colspan="2" class="first" style="background:#FFC">��������</th>
								<td class="left">
                                <select name="app_hm_type" id="app_hm_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_hm_type = "" then %>selected<% end if %>>����</option>
				                    <option value='����' <%If app_hm_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='¡��' <%If app_hm_type = "¡��" then %>selected<% end if %>>¡��</option>
                                    <option value='��Ÿ' <%If app_hm_type = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
								<th style="background:#FFC">���� Comment.</th>
								<td colspan="8" class="left">
								<input name="app_hm_comment" type="text" id="app_hm_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--�����߷�-->
                            <tr style="display:none;" id="hu_menu1">
								<th colspan="2" class="first" style="background:#FFC">��������</th>
								<td colspan="3" class="left">
                                <select name="app_hu_type" id="app_hu_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_hu_type = "" then %>selected<% end if %>>����</option>
				                    <option value='����' <%If app_hu_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='����' <%If app_hu_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='����' <%If app_hu_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='����' <%If app_hu_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='���λ���' <%If app_hu_type = "���λ���" then %>selected<% end if %>>���λ���</option>
                                </select>
                                <th style="background:#FFC">�����Ⱓ</th>
								<td colspan="6" class="left">
                                <input name="app_hustart_date" type="text" size="10" readonly="true" id="datepicker2" style="width:70px;">
                                &nbsp;&nbsp;��&nbsp;&nbsp;
                                <input name="app_hufinish_date" type="text" size="10" readonly="true" id="datepicker3" style="width:70px;">&nbsp;</td>
                            </tr>
                            <tr style="display:none;" id="hu_menu2">
								<th colspan="2" class="first" style="background:#FFC">���� Comment.</th>
								<td colspan="10" class="left">
								<input name="app_hu_comment" type="text" id="app_hu_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--¡��߷�-->
							<%
							Dim app_di_type
							%>
                            <tr style="display:none;" id="di_menu1">
								<th colspan="2" class="first" style="background:#FFC">¡������</th>
								<td colspan="3" class="left">
                                <select name="app_di_type" id="app_di_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_di_type = "" then %>selected<% end if %>>����</option>
				                    <option value='���߷�' <%If app_di_type = "���߷�" then %>selected<% end if %>>���߷�</option>
                                    <option value='��������' <%If app_di_type = "��������" then %>selected<% end if %>>��������</option>
                                    <option value='����' <%If app_di_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='����' <%If app_di_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='�ư�' <%If app_di_type = "�ư�" then %>selected<% end if %>>�ư�</option>
                                </select>
                                <th style="background:#FFC">¡��Ⱓ</th>
								<td colspan="6" class="left">
                                <input name="app_distart_date" type="text" size="10" readonly="true" id="datepicker4" style="width:70px;">
                                &nbsp;&nbsp;��&nbsp;&nbsp;
                                <input name="app_difinish_date" type="text" size="10" readonly="true" id="datepicker5" style="width:70px;">&nbsp;</td>
                            </tr>
                            <tr style="display:none;" id="di_menu2">
								<th colspan="2" class="first" style="background:#FFC">¡�� Comment.</th>
								<td colspan="10" class="left">
								<input name="app_di_comment" type="text" id="app_di_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--���� �u��-->
							<%
							Dim app_rw_type
							%>
                            <tr style="display:none;" id="rw_menu1">
								<th colspan="2" class="first" style="background:#FFC">��������</th>
								<td colspan="3" class="left">
                                <select name="app_rw_type" id="app_rw_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_rw_type = "" then %>selected<% end if %>>����</option>
				                    <option value='Ư������' <%If app_rw_type = "Ư������" then %>selected<% end if %>>Ư������</option>
                                    <option value='��������' <%If app_rw_type = "��������" then %>selected<% end if %>>��������</option>
                                </select>
								<th style="background:#FFC">���� Comment.</th>
								<td colspan="7" class="left">
								<input name="app_rw_comment" type="text" id="app_rw_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                </div>
                <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="emp_name" value="<%=emp_name%>" ID="Hidden1">
                <input type="hidden" name="app_grade" value="<%=emp_grade%>" ID="Hidden1">
                <input type="hidden" name="app_position" value="<%=emp_position%>" ID="Hidden1">
                <input type="hidden" name="app_job" value="<%=emp_job%>" ID="Hidden1">
                <input type="hidden" name="app_to_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="app_to_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="app_to_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="app_to_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="app_org" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="app_org_name" value="<%=emp_org_name%>" ID="Hidden1">

                <input type="hidden" name="t_date" value="<%=t_date%>" ID="Hidden1">
        	</form>
		</div>
	</div>
	</body>
</html>

