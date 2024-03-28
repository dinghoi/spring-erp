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

title_line = " 인사 발령 등록 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
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
					alert('발령일을 입력하세요');
					frm.app_date.focus();
					return false;}

//				if(document.frm.app_id.value !="퇴직발령")
//				    if(document.frm.app_date.value < document.frm.t_date.value) {
//					    alert('발령일이 전월입니다. 확인하십시요!');
//					    frm.app_date.focus();
//					    return false;}

				if(document.frm.app_id.value =="") {
					alert('발령구분을 선택하세요');
					frm.app_id.focus();
					return false;}

				if(document.frm.app_id.value =="이동발령")
					if(document.frm.app_be_orgcode.value =="") {
						alert('발령소속을 입력하세요');
						frm.app_be_orgcode.focus();
						return false;}
				if(document.frm.app_id.value =="이동발령")
					if(document.frm.app_mv_comment.value =="") {
						alert('발령사유를 입력하세요');
						frm.app_mv_comment.focus();
						return false;}

				if(document.frm.app_id.value =="퇴직발령")
					if(document.frm.app_date.value =="") {
						alert('퇴직일을 입력하세요');
						frm.app_date.focus();
						return false;}
				if(document.frm.app_id.value =="퇴직발령")
					if(document.frm.app_end_type.value =="") {
						alert('퇴직유형을 선택하세요');
						frm.app_end_type.focus();
						return false;}

				if(document.frm.app_id.value =="승진발령")
					if(document.frm.app_be_grade.value =="") {
						alert('승진직급을 선택하세요');
						frm.app_be_grade.focus();
						return false;}
				if(document.frm.app_id.value =="승진발령")
					if(document.frm.app_gr_type.value =="") {
						alert('승진유형을 선택하세요');
						frm.app_gr_type.focus();
						return false;}

				if(document.frm.app_id.value =="직책보임")
					if(document.frm.app_be_position.value =="") {
						alert('보임직책을 선택하세요');
						frm.app_be_position.focus();
						return false;}
				if(document.frm.app_id.value =="직책보임")
					if(document.frm.app_bm_orgcode.value =="") {
						alert('보임소속을 입력하세요');
						frm.app_bm_orgcode.focus();
						return false;}

                if(document.frm.app_id.value =="직책해임")
					if(document.frm.app_hm_type.value =="") {
						alert('해임유형을 선택하세요');
						frm.app_hm_type.focus();
						return false;}

				if(document.frm.app_id.value =="휴직발령")
					if(document.frm.app_hu_type.value =="") {
						alert('휴직유형을 선택하세요');
						frm.app_hu_type.focus();
						return false;}
				if(document.frm.app_id.value =="휴직발령")
					if(document.frm.app_hustart_date.value =="") {
						alert('휴직시작일을 입력하세요');
						frm.app_hustart_date.focus();
						return false;}
				if(document.frm.app_id.value =="휴직발령")
					if(document.frm.app_hufinish_date.value =="") {
						alert('휴직만료일을 입력하세요');
						frm.app_hufinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="휴직발령")
				    if(document.frm.app_hustart_date.value > document.frm.app_hufinish_date.value) {
						alert('휴직시작일이 휴직만료일보다 늦습니다');
						frm.app_hufinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="휴직발령")
				    if(document.frm.app_date.value > document.frm.app_hustart_date.value) {
						alert('발령일보다 휴직시작일이 빠름니다');
						frm.app_hustart_date.focus();
						return false;}

				if(document.frm.app_id.value =="징계발령")
					if(document.frm.app_di_type.value =="") {
						alert('징계유형을 선택하세요');
						frm.app_di_type.focus();
						return false;}
				if(document.frm.app_id.value =="징계발령")
					if(document.frm.app_distart_date.value =="") {
						alert('징계시작일을 입력하세요');
						frm.app_distart_date.focus();
						return false;}
				if(document.frm.app_id.value =="징계발령")
					if(document.frm.app_difinish_date.value =="") {
						alert('징계만료일을 입력하세요');
						frm.app_difinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="징계발령")
				    if(document.frm.app_distart_date.value > document.frm.app_difinish_date.value) {
						alert('징계시작일이 징계만료일보다 늦습니다');
						frm.app_hufinish_date.focus();
						return false;}
				if(document.frm.app_id.value =="징계발령")
				    if(document.frm.app_date.value > document.frm.app_distart_date.value) {
						alert('발령일보다 징계시작일이 빠름니다');
						frm.app_distart_date.focus();
						return false;}

				if(document.frm.app_id.value =="포상발령")
					if(document.frm.app_rw_type.value =="") {
						alert('포상유형을 선택하세요');
						frm.app_rw_type.focus();
						return false;}
				if(document.frm.app_id.value =="포상발령")
					if(document.frm.app_rw_comment.value =="") {
						alert('포상내용을 입력하세요');
						frm.app_rw_comment.focus();
						return false;}

				{
				a=confirm('발령처리를 하시겠습니까?')
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
				if (c == '이동발령')
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
				if (c == '퇴직발령')
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
				if (c == '승진발령')
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
				if (c == '직책보임')
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
				if (c == '직책해임')
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
				if (c == '휴직발령')
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
				if (c == '징계발령')
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
				if (c == '포상발령')
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
								<th>사원&nbsp;&nbsp;번호</th>
                                <td class="left"><%=emp_no%>&nbsp;</td>
                                <th>성명(한글)</th>
                                <td class="left"><%=emp_name%>&nbsp;</td>
								<th>성명(영문)</th>
								<td colspan="2" class="left"><%=emp_ename%>&nbsp;</td>
                                <th>생년월일</th>
                                <td colspan="2" class="left"><%=emp_birthday%>&nbsp;(<%=emp_birthday_id%>)&nbsp;</td>
                           </tr>
                           <tr>
                                <th>소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;―&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th>조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th>직원구분</th>
                                <td class="left"><%=emp_type%>&nbsp;</td>
                               	<th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;급</th>
								<td class="left"><%=emp_grade%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;위</th>
								<td class="left"><%=emp_job%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;책</th>
                                <td class="left"><%=emp_position%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;무</th>
								<td class="left"><%=emp_jikmu%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th>최초입사일</th>
                                <td class="left"><%=emp_first_date%>&nbsp;</td>
                                <th>입&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;일</th>
                                <td class="left"><%=emp_in_date%>&nbsp;</td>
                                <th>퇴직기산일</th>
                                <td class="left"><%=emp_end_gisan%>&nbsp;</td>
                                <th>근속기산일</th>
                                <td class="left"><%=emp_gunsok_date%>&nbsp;</td>
                                <th>연차기산일</th>
                                <td class="left"><%=emp_yuncha_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">주민번호</th>
								<td colspan="2" class="left"><%=emp_person1%>&nbsp;―&nbsp;<%=emp_person2%>&nbsp;(<%=emp_sex%>)&nbsp;</td>
                                <th>전화번호</th>
								<td colspan="3" class="left"><%=emp_tel_ddd%>―<%=emp_tel_no1%>―<%=emp_tel_no2%>&nbsp;</td>
                                <th>핸드폰</th>
								<td colspan="3" class="left"><%=emp_hp_ddd%>―<%=emp_hp_no1%>―<%=emp_hp_no2%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">현소속발령일</th>
								<td class="left"><%=emp_org_baldate%>&nbsp;</td>
                                <th>현직급승진일</th>
								<td class="left"><%=emp_grade_date%>&nbsp;</td>
                                <th>e_메일주소</th>
								<td colspan="2" class="left"><%=emp_email%>&nbsp;</td>
                                <th>경조가입</th>
								<td colspan="3" class="left">
                                <input type="radio" name="emp_sawo_id" value="Y" <% if emp_sawo_id = "Y" then %>checked<% end if %>>가입
              					<input name="emp_sawo_id" type="radio" value="N" <% if emp_sawo_id = "N" then %>checked<% end if %>>안함
                                 &nbsp;&nbsp;―&nbsp;&nbsp;<%=emp_sawo_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="12" class="left" style="background:#FFC">■ 인사&nbsp;&nbsp;&nbsp;발령 ■</th>&nbsp;
                            </tr>
                            <tr>
                                <th colspan="2" class="first">인사발령일자</th>
                                <td colspan="3" class="left">
                                <input name="app_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;">&nbsp;</td>
                                <th class="first">발령구분</th>&nbsp;
                                <td colspan="6" class="left">
                            <%
								'Sql="select * from emp_etc_code where emp_etc_type = '10' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '10' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							%>
								<select name="app_id" id="select" style="width:150px" onChange="menu1()">
                                <option value="" <% if app_id = "" then %>selected<% end if %>>선택</option>
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

							<!--이동 발령-->
							<tr style="display:none;" id="mv_menu1">
								<th colspan="2" class="first" >현소속</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;―&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th class="first" >현조직</th>
                                <td colspan="6" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
							<%
							Dim app_be_orgcode, app_company, app_bonbu, app_saupbu, app_team
							Dim app_reside_place, emp_org_level, app_reside_company
							%>
                            <tr style="display:none;" id="mv_menu2">
								<th colspan="2" class="first" style="background:#FFC">발령소속</th>
								<td colspan="3" class="left">
								<input name="app_be_orgcode" type="text" id="app_be_orgcode" style="width:40px" readonly="true" value="<%=app_be_orgcode%>">
                                &nbsp;―&nbsp;
                                <input name="app_be_org" type="text" id="app_be_org" style="width:120px" readonly="true" value="<%=app_be_org%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="apporg"%>&view_condi=<%=emp_company%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
                                </td>
                                <th style="background:#FFC">발령조직</th>
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
                                <th colspan="2" class="first" style="background:#FFC">실근무지/주소</th>
                                <td colspan="3" class="left">
                                <input name="emp_stay_code" type="text" id="emp_stay_code" style="width:40px" readonly="true" value="<%=emp_stay_code%>">
                                &nbsp;―&nbsp;
                                <input name="stay_name" type="text" id="stay_name" style="width:150px"  value="<%=stay_name%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_stay_select.asp?gubun=<%="stay"%>&reside_code=<%=emp_stay_code%>','stayselect','scrollbars=yes,width=1000,height=400')">선택</a>
                                </td>
                                <th>주소지</th>
                                <td colspan="6" class="left">
                                <input name="stay_sido" type="text" id="stay_sido" style="width:100px" readonly="true" value="<%=stay_sido%>">
                                <input name="stay_gugun" type="text" id="stay_gugun" style="width:150px" readonly="true" value="<%=stay_gugun%>">
                                <input name="stay_dong" type="text" id="stay_dong" style="width:150px" readonly="true" value="<%=stay_dong%>">
                                <input name="stay_addr" type="text" id="stay_addr" style="width:200px" readonly="true" value="<%=stay_addr%>">
								</td>
                            </tr>
                            <tr style="display:none;" id="mv_menu4">
                                <th colspan="2" class="first" style="background:#FFC">직무</th>
                                <td class="left">
                                <%
								'Sql="select * from emp_etc_code where emp_etc_type = '05' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '05' ORDER BY emp_etc_code ASC"

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
								<select name="emp_jikmu" id="emp_jikmu" style="width:90px">
                                <option value="" <% if emp_jikmu = "" then %>selected<% end if %>>선택</option>
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
                                <th>비용구분</th>
                                <td class="left">
                              <%
								'Sql="select * from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
								<select name="cost_center" id="cost_center" style="width:90px">
                                <option value="" <% if cost_center = "" then %>selected<% end if %>>선택</option>
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
								<th style="background:#FFC">발령사유</th>
								<td colspan="6" class="left">
								<input name="app_mv_comment" type="text" id="app_mv_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--퇴직 발령-->
							<%
							Dim app_end_type
							%>
                            <tr style="display:none;" id="end_menu1">
                                <th colspan="2" style="background:#FFC">퇴직유형</th>
								<td colspan="2" class="left">
                                <select name="app_end_type" id="app_end_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_end_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='회사사정' <%If app_end_type = "회사사정" then %>selected<% end if %>>회사사정</option>
                                    <option value='명예퇴직' <%If app_end_type = "명예퇴직" then %>selected<% end if %>>명예퇴직</option>
                                    <option value='개인사정' <%If app_end_type = "개인사정" then %>selected<% end if %>>개인사정</option>
                                    <option value='징계' <%If app_end_type = "징계" then %>selected<% end if %>>징계</option>
                                    <option value='육아' <%If app_end_type = "육아" then %>selected<% end if %>>육아</option>
                                    <option value='간병' <%If app_end_type = "간병" then %>selected<% end if %>>간병</option>
                                    <option value='치료' <%If app_end_type = "치료" then %>selected<% end if %>>치료</option>
                                </select>
								<th style="background:#FFC">퇴직 Comment.</th>
								<td colspan="7" class="left">
								<input name="app_end_comment" type="text" id="app_end_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--승진 발령-->
							<%
							Dim app_gr_type
							%>
							<tr style="display:none;" id="gr_menu1">
								<th colspan="2" class="first" >현직급</th>
								<td class="left"><%=emp_grade%>&nbsp;</td>
                                <th>현직위</th>
								<td class="left"><%=emp_job%>&nbsp;</td>
                                <th>현직급 승진일</th>
                                <td colspan="6" class="left"><%=emp_grade_date%>&nbsp;</td>
							</tr>
                            <tr style="display:none;" id="gr_menu2">
								<th colspan="2" class="first" style="background:#FFC">승진직급</th>
								<td class="left">
                            <%
								'Sql="select * from emp_etc_code where emp_etc_type = '02' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '02' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							%>
								<select name="app_be_grade" id="app_be_grade" style="width:90px">
                                <option value="" <% if app_be_grade = "" then %>selected<% end if %>>선택</option>
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
                                <th style="background:#FFC">승진유형</th>
								<td class="left">
                                <select name="app_gr_type" id="app_gr_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_gr_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='정기승진' <%If app_gr_type = "정기승진" then %>selected<% end if %>>정기승진</option>
                                    <option value='특별승진' <%If app_gr_type = "특별승진" then %>selected<% end if %>>특별승진</option>
                                    <option value='직권승진' <%If app_gr_type = "직권승진" then %>selected<% end if %>>직권승진</option>
                                </select>
								<th style="background:#FFC">승진 Comment.</th>
								<td colspan="6" class="left">
								<input name="app_gr_comment" type="text" id="app_gr_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--직책보임-->
							<%
							Dim app_be_position, app_bm_orgcode, app_bm_org, app_bm_company, app_bm_bonbu
							Dim app_bm_saupbu, app_bm_team, app_bm_reside_place, app_bm_reside_company
							Dim app_bm_org_level
							%>
							<tr style="display:none;" id="bm_menu1">
                                <th colspan="2" class="first">현&nbsp;직책</th>
								<td class="left"><%=emp_position%>&nbsp;</td>
								<th >현&nbsp;소속</th>
								<td colspan="2" class="left"><%=emp_org_code%>&nbsp;―&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th >현&nbsp;조직</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
                            <tr style="display:none;" id="bm_menu2">
								<th colspan="2" class="first" style="background:#FFC">보임직책</th>
								<td class="left">
                              <%
								'Sql="select * from emp_etc_code where emp_etc_type = '04' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '04' ORDER BY emp_etc_code ASC "

								Rs_etc.Open objBuilder.ToString(), DBConn, 1
							  %>
								<select name="app_be_position" id="app_be_position" style="width:90px">
                                <option value="" <% if app_be_position = "" then %>selected<% end if %>>선택</option>
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
								<th style="background:#FFC">보임소속</th>
								<td colspan="2" class="left">
								<input name="app_bm_orgcode" type="text" id="app_bm_orgcode" style="width:30px" readonly="true" value="<%=app_bm_orgcode%>">
                                ―
                                <input name="app_bm_org" type="text" id="app_bm_org" style="width:80px" readonly="true" value="<%=app_bm_org%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="appbmorg"%>&view_condi=<%=emp_company%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
                                </td>
                                <th style="background:#FFC">보임조직</th>
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
								<th colspan="2" class="first" style="background:#FFC">보임 Comment.</th>
								<td colspan="10" class="left">
								<input name="app_bm_comment" type="text" id="app_bm_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--직책해임-->
							<%
							Dim app_hm_type, app_hu_type
							%>
							<tr style="display:none;" id="hm_menu1">
                                <th colspan="2" class="first">현&nbsp;직책</th>
								<td class="left"><%=emp_position%>&nbsp;</td>
								<th >현&nbsp;소속</th>
								<td colspan="2" class="left"><%=emp_org_code%>&nbsp;―&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th >현&nbsp;조직</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
                            <tr style="display:none;" id="hm_menu2">
								<th colspan="2" class="first" style="background:#FFC">해임유형</th>
								<td class="left">
                                <select name="app_hm_type" id="app_hm_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_hm_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='퇴직' <%If app_hm_type = "퇴직" then %>selected<% end if %>>퇴직</option>
                                    <option value='징계' <%If app_hm_type = "징계" then %>selected<% end if %>>징계</option>
                                    <option value='기타' <%If app_hm_type = "기타" then %>selected<% end if %>>기타</option>
                                </select>
								<th style="background:#FFC">해임 Comment.</th>
								<td colspan="8" class="left">
								<input name="app_hm_comment" type="text" id="app_hm_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--휴직발령-->
                            <tr style="display:none;" id="hu_menu1">
								<th colspan="2" class="first" style="background:#FFC">휴직유형</th>
								<td colspan="3" class="left">
                                <select name="app_hu_type" id="app_hu_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_hu_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='병가' <%If app_hu_type = "병가" then %>selected<% end if %>>병가</option>
                                    <option value='육아' <%If app_hu_type = "육아" then %>selected<% end if %>>육아</option>
                                    <option value='간병' <%If app_hu_type = "간병" then %>selected<% end if %>>간병</option>
                                    <option value='가사' <%If app_hu_type = "가사" then %>selected<% end if %>>가사</option>
                                    <option value='개인사정' <%If app_hu_type = "개인사정" then %>selected<% end if %>>개인사정</option>
                                </select>
                                <th style="background:#FFC">휴직기간</th>
								<td colspan="6" class="left">
                                <input name="app_hustart_date" type="text" size="10" readonly="true" id="datepicker2" style="width:70px;">
                                &nbsp;&nbsp;∼&nbsp;&nbsp;
                                <input name="app_hufinish_date" type="text" size="10" readonly="true" id="datepicker3" style="width:70px;">&nbsp;</td>
                            </tr>
                            <tr style="display:none;" id="hu_menu2">
								<th colspan="2" class="first" style="background:#FFC">휴직 Comment.</th>
								<td colspan="10" class="left">
								<input name="app_hu_comment" type="text" id="app_hu_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--징계발령-->
							<%
							Dim app_di_type
							%>
                            <tr style="display:none;" id="di_menu1">
								<th colspan="2" class="first" style="background:#FFC">징계유형</th>
								<td colspan="3" class="left">
                                <select name="app_di_type" id="app_di_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_di_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='대기발령' <%If app_di_type = "대기발령" then %>selected<% end if %>>대기발령</option>
                                    <option value='직무정지' <%If app_di_type = "직무정지" then %>selected<% end if %>>직무정지</option>
                                    <option value='감봉' <%If app_di_type = "감봉" then %>selected<% end if %>>감봉</option>
                                    <option value='강등' <%If app_di_type = "강등" then %>selected<% end if %>>강등</option>
                                    <option value='훈계' <%If app_di_type = "훈계" then %>selected<% end if %>>훈계</option>
                                </select>
                                <th style="background:#FFC">징계기간</th>
								<td colspan="6" class="left">
                                <input name="app_distart_date" type="text" size="10" readonly="true" id="datepicker4" style="width:70px;">
                                &nbsp;&nbsp;∼&nbsp;&nbsp;
                                <input name="app_difinish_date" type="text" size="10" readonly="true" id="datepicker5" style="width:70px;">&nbsp;</td>
                            </tr>
                            <tr style="display:none;" id="di_menu2">
								<th colspan="2" class="first" style="background:#FFC">징계 Comment.</th>
								<td colspan="10" class="left">
								<input name="app_di_comment" type="text" id="app_di_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>

							<!--포상 뱔령-->
							<%
							Dim app_rw_type
							%>
                            <tr style="display:none;" id="rw_menu1">
								<th colspan="2" class="first" style="background:#FFC">포상유형</th>
								<td colspan="3" class="left">
                                <select name="app_rw_type" id="app_rw_type" value="<%=app_id_type%>" style="width:80px">
			            	        <option value="" <% if app_rw_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='특별포상' <%If app_rw_type = "특별포상" then %>selected<% end if %>>특별포상</option>
                                    <option value='정기포상' <%If app_rw_type = "정기포상" then %>selected<% end if %>>정기포상</option>
                                </select>
								<th style="background:#FFC">포상 Comment.</th>
								<td colspan="7" class="left">
								<input name="app_rw_comment" type="text" id="app_rw_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();"></span>
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

