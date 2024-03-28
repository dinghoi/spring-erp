<!--#include virtual="/common/inc_top_utf.asp"-->
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
Dim sch_tab(10, 10)
Dim car_tab(20, 10)
Dim qul_tab(20, 10)
Dim fam_tab(10, 10)
Dim app_tab(50, 30)
Dim edu_tab(10, 10)
Dim lan_tab(10, 10)

Dim curr_date, be_pg, page, view_sord, page_cnt, view_sort
Dim photo_image, emp_email, emp_person2, emp_end_date
Dim emp_grade_date, emp_org_baldate, emp_marry_date
Dim emp_military_date1, emp_military_date2, i, j, k
Dim rs, rs_sch, k_sch, rs_car, k_car
Dim rs_qul, k_qul, rs_fam, k_fam
Dim rs_app, k_app, rs_edu, k_edu, rs_lan, k_lan
Dim stay_name, stay_sido, stay_gugun, stay_dong, stay_addr
Dim stay_code, sawo_id

emp_no = f_Request("emp_no")

curr_date = Mid(CStr(Now()), 1, 10)

objBuilder.Append "SELECT emtt.emp_image, emtt.emp_end_date, emtt.emp_grade_date, emtt.emp_org_baldate, "
objBuilder.Append "	emtt.emp_marry_date, emtt.emp_military_date1, emtt.emp_military_date2, emtt.emp_no, emtt.emp_name, "
objBuilder.Append "	emtt.emp_ename, emtt.emp_org_code, emtt.emp_org_name, emtt.emp_grade, emtt.emp_job, "
objBuilder.Append "	emtt.emp_position, emtt.emp_in_date, emtt.emp_person1, emtt.emp_person2, emtt.emp_birthday, "
objBuilder.Append "	emtt.emp_sex, emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_hp_ddd, "
objBuilder.Append "	emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, "
objBuilder.Append "	emtt.emp_emergency_tel, emtt.emp_family_sido, emtt.emp_family_gugun, emtt.emp_family_dong, "
objBuilder.Append "	emtt.emp_family_addr, emtt.emp_military_id, emtt.emp_military_grade, emtt.emp_military_comm, "
objBuilder.Append "	emtt.emp_faith, emtt.emp_hobby, emtt.emp_sawo_id, emtt.emp_disabled, emtt.emp_disab_grade, "
objBuilder.Append "	emtt.emp_first_date, emtt.emp_gunsok_date, emtt.emp_end_gisan, emtt.emp_email, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name, emtt.emp_stay_code "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"' "

Set rs = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

photo_image = "/emp_photo/" & rs("emp_image")
emp_email = rs("emp_email") & "@k-one.co.kr"
emp_person2 = "*******"

'입력받지 못하는 날짜필드를 처음 1900-01-01로 하놔서..ㅠㅠ
If rs("emp_end_date") = "1900-01-01" Then
   emp_end_date = ""
Else
   emp_end_date = rs("emp_end_date")
End If

If rs("emp_grade_date") = "1900-01-01" Then
   emp_grade_date = ""
Else
   emp_grade_date = rs("emp_grade_date")
End If

If rs("emp_org_baldate") = "1900-01-01" Then
   emp_org_baldate = ""
Else
   emp_org_baldate = rs("emp_org_baldate")
End If

If rs("emp_marry_date") = "1900-01-01" Then
   emp_marry_date = ""
Else
   emp_marry_date = rs("emp_marry_date")
End If

If rs("emp_military_date1") = "1900-01-01" Then
   emp_military_date1 = ""
Else
   emp_military_date1 = rs("emp_military_date1")
End If

If rs("emp_military_date2") = "1900-01-01" Then
   emp_military_date2 = ""
Else
   emp_military_date2 = rs("emp_military_date2")
End If

'학력사항 db
For i = 0 To 10
'	com_tab(i) = ""
'	com_sum(i) = 0
	For j = 0 To 10
		sch_tab(i, j) = ""
'		com_in(i,j) = 0
	Next
Next

k = 0

objBuilder.Append "SELECT sch_start_date, sch_end_date, sch_school_name, sch_dept, "
objBuilder.Append "	sch_major, sch_sub_major, sch_degree, sch_finish "
objBuilder.Append "FROM emp_school "
objBuilder.Append "WHERE sch_empno = '"&emp_no&"' "
objBuilder.Append "ORDER BY sch_empno, sch_seq ASC "

Set rs_sch = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_sch.EOF
	k = k + 1

	sch_tab(k, 1) = rs_sch("sch_start_date")
	sch_tab(k, 2) = rs_sch("sch_end_date")
	sch_tab(k, 3) = rs_sch("sch_school_name")
	sch_tab(k, 4) = rs_sch("sch_dept")
	sch_tab(k, 5) = rs_sch("sch_major")
	sch_tab(k, 6) = rs_sch("sch_sub_major")
	sch_tab(k, 7) = rs_sch("sch_degree")
	sch_tab(k, 8) = rs_sch("sch_finish")

	rs_sch.MoveNext()
Wend
rs_sch.Close() : Set rs_sch = Nothing

k_sch = k

'경력사항 db
For i = 0 To  20
	For j = 0 To  10
		car_tab(i, j) = ""
	Next
Next

k = 0

objBuilder.Append "SELECT career_join_date, career_end_date, career_office, "
objBuilder.Append "	career_dept, career_position, career_task "
objBuilder.Append "FROM emp_career "
objBuilder.Append "WHERE career_empno = '"&emp_no&"' "
objBuilder.Append "ORDER BY career_empno, career_seq ASC "

Set rs_car = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_car.EOF
	k = k + 1

	car_tab(k, 1) = rs_car("career_join_date")
	car_tab(k, 2) = rs_car("career_end_date")
	car_tab(k, 3) = rs_car("career_office")
	car_tab(k, 4) = rs_car("career_dept")
	car_tab(k, 5) = rs_car("career_position")
	car_tab(k, 6) = rs_car("career_task")

	rs_car.MoveNext()
Wend
rs_car.Close() : Set rs_car = Nothing

k_car = k

'자격사항 db
For i = 0 To 20
	For j = 0 To 10
		qul_tab(i, j) = ""
	Next
Next

k = 0

objBuilder.Append "SELECT qual_type, qual_grade, qual_pass_date, qual_org, qual_no "
objBuilder.Append "FROM emp_qual "
objBuilder.Append "WHERE qual_empno = '"&emp_no&"' "
objBuilder.Append "ORDER BY qual_empno, qual_seq ASC "

Set rs_qul = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_qul.EOF
	k = k + 1

	qul_tab(k, 1) = rs_qul("qual_type")
	qul_tab(k, 2) = rs_qul("qual_grade")
	qul_tab(k, 3) = rs_qul("qual_pass_date")
	qul_tab(k, 4) = rs_qul("qual_org")
	qul_tab(k, 5) = rs_qul("qual_no")

	rs_qul.MoveNext()
Wend
rs_qul.Close() : Set rs_qul = Nothing

k_qul = k

'가족사항 db
For i = 0 To 10
	For j = 0 To 10
		fam_tab(i, j) = ""
	Next
Next

k = 0

objBuilder.Append "SELECT family_rel, family_name, family_birthday, family_birthday_id, family_job, "
objBuilder.Append "	family_tel_ddd, family_tel_no1, family_tel_no2, family_live, family_person1, family_person2 "
objBuilder.Append "FROM emp_family "
objBuilder.Append "WHERE family_empno = '"&emp_no&"' "
objBuilder.Append "ORDER BY family_empno, family_seq ASC "

Set rs_fam = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_fam.EOF
	k = k + 1

	fam_tab(k, 1) = rs_fam("family_rel")
	fam_tab(k, 2) = rs_fam("family_name")
	fam_tab(k, 3) = rs_fam("family_birthday")
	fam_tab(k, 4) = rs_fam("family_birthday_id")
	fam_tab(k, 5) = rs_fam("family_job")
	fam_tab(k, 6) = rs_fam("family_tel_ddd") & "-" & rs_fam("family_tel_no1") & "-" & rs_fam("family_tel_no2")
	fam_tab(k, 7) = rs_fam("family_live")
	fam_tab(k, 8) = rs_fam("family_person1")
	fam_tab(k, 9) = rs_fam("family_person2")

	rs_fam.MoveNext()
Wend
rs_fam.close() : Set rs_fam = Nothing

k_fam = k

'발령사항 db
For i = 0 To 50
	For j = 0 To 30
		app_tab(i, j) = ""
	Next
Next

k = 0

objBuilder.Append "SELECT app_date, app_id, app_id_type, app_to_company, app_to_orgcode, app_to_org, "
objBuilder.Append "	app_to_grade, app_to_job, app_to_position, app_to_enddate, app_be_company, "
objBuilder.Append "	app_be_orgcode, app_be_org, app_be_grade, app_be_job, app_be_position, app_be_enddate, "
objBuilder.Append "	app_start_date, app_finish_date, app_reward, app_comment "
objBuilder.Append "FROM emp_appoint "
objBuilder.Append "WHERE app_empno = '"&emp_no&"' "
objBuilder.Append "ORDER BY app_empno, app_seq ASC "

Set rs_app = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_app.EOF
	k = k + 1

	app_tab(k, 1) = rs_app("app_date")
	app_tab(k, 2) = rs_app("app_id")
	app_tab(k, 3) = rs_app("app_id_type")
	app_tab(k, 4) = rs_app("app_to_company")
	app_tab(k, 5) = rs_app("app_to_orgcode")
	app_tab(k, 6) = rs_app("app_to_org")
	app_tab(k, 7) = rs_app("app_to_grade")
	app_tab(k, 8) = rs_app("app_to_job")
	app_tab(k, 9) = rs_app("app_to_position")
	app_tab(k, 10) = rs_app("app_to_enddate")
	app_tab(k, 11) = rs_app("app_be_company")
	app_tab(k, 12) = rs_app("app_be_orgcode")
	app_tab(k, 13) = rs_app("app_be_org")
	app_tab(k, 14) = rs_app("app_be_grade")
	app_tab(k, 15) = rs_app("app_be_job")
	app_tab(k, 16) = rs_app("app_be_position")
	app_tab(k, 17) = rs_app("app_be_enddate")
	app_tab(k, 18) = rs_app("app_start_date")
	app_tab(k, 19) = rs_app("app_finish_date")
	app_tab(k, 20) = rs_app("app_reward")
	app_tab(k, 21) = rs_app("app_comment")

	rs_app.MoveNext()
Wend
rs_app.Close() : Set rs_app = Nothing

k_app = k

'교육사항 db
For i = 0 To 10
	For j = 0 To 10
		edu_tab(i, j) = ""
	Next
Next

k = 0

objBuilder.Append "SELECT edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date, edu_comment "
objBuilder.Append "FROM emp_edu "
objBuilder.Append "WHERE edu_empno = '"&emp_no&"' "
objBuilder.Append "ORDER BY edu_empno, edu_seq ASC "

Set rs_edu = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_edu.EOF
	k = k + 1

	edu_tab(k, 1) = rs_edu("edu_name")
	edu_tab(k, 2) = rs_edu("edu_office")
	edu_tab(k, 3) = rs_edu("edu_finish_no")
	edu_tab(k, 4) = rs_edu("edu_start_date")
	edu_tab(k, 5) = rs_edu("edu_end_date")
	edu_tab(k, 6) = rs_edu("edu_comment")

	rs_edu.MoveNext()
Wend
rs_edu.Close() : Set rs_edu = Nothing

k_edu = k

'어학사항 db
For i = 0 To 10
	For j = 0 To 10
		lan_tab(i, j) = ""
	Next
Next

k = 0

objbuilder.Append "SELECT lang_id, lang_id_type, lang_point, lang_grade, lang_get_date "
objbuilder.Append "FROM emp_language "
objbuilder.Append "WHERE lang_empno = '"&emp_no&"' "
objbuilder.Append "ORDER BY lang_empno, lang_seq ASC "

Set rs_lan = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

While Not rs_lan.EOF
	k = k + 1

	lan_tab(k, 1) = rs_lan("lang_id")
	lan_tab(k, 2) = rs_lan("lang_id_type")
	lan_tab(k, 3) = rs_lan("lang_point")
	lan_tab(k, 4) = rs_lan("lang_grade")
	lan_tab(k, 5) = rs_lan("lang_get_date")

	rs_lan.MoveNext()
Wend
rs_lan.Close() : Set rs_lan = Nothing

k_lan = k

'실근무지주소
stay_name = ""
stay_sido = ""
stay_gugun = ""
stay_dong = " "
stay_addr = ""
stay_code = rs("emp_stay_code")

If stay_code <> "" Then
	objBuilder.Append "SELECT stay_name, stay_sido, stay_gugun, stay_dong, stay_addr "
	objBuilder.Append "FROM emp_stay "
	objBuilder.Append "WHERE stay_code = '"&stay_code&"' "

	Set rs_stay = DBConn.Execute(objBuilder.ToString())

	If Not rs_stay.EOF Then
		stay_name = rs_stay("stay_name")
		stay_sido = rs_stay("stay_sido")
		stay_gugun = rs_stay("stay_gugun")
		stay_dong = rs_stay("stay_dong")
		stay_addr = rs_stay("stay_addr")
	End If
	rs_stay.Close() : Set rs_stay = Nothing
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>인사 관리 시스템</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script src="/java/common.js" type="text/javascript"></script>

<script type="text/javascript">
	function printWindow(){
//		viewOff("button");
		factory.printing.header = ""; //머리말 정의
		factory.printing.footer = ""; //꼬리말 정의
		factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
		factory.printing.leftMargin = 13; //외쪽 여백 설정
		factory.printing.topMargin = 25; //윗쪽 여백 설정
		factory.printing.rightMargin = 13; //오른쯕 여백 설정
		factory.printing.bottomMargin = 15; //바닦 여백 설정
//		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
//		factory.printing.printer = ""; //프린터 할 프린터 이름
//		factory.printing.paperSize = "A4"; //용지선택
//		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
//		factory.printing.collate = true; //순서대로 출력하기
//		factory.printing.copies = "1"; //인쇄할 매수
//		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
//		factory.printing.Printer(true); //출력하기
		factory.printing.Preview(); //윈도우를 통해서 출력
		factory.printing.Print(false); //윈도우를 통해서 출력
	}

	function printW(){
        window.print();
    }

	function goBefore(){
		history.back();
	}

	//프린트 함수 신규 작성[허정호_20220204]
	var printArea;
	var initBody;

	function fnPrint(id){
		printArea = document.getElementById(id);

		window.onbeforeprint = beforePrint;
		window.onafterprint = afterPrint;

		window.print();
	}

	function beforePrint(){
		initBody = document.body.innerHTML;
		document.body.innerHTML = printArea.innerHTML;
	}

	function afterPrint(){
		document.body.innerHTML = initBody;
	}
</script>

<style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
</style>
<style media="print">
	.noprint{ display: none }
</style>
</head>

<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<div class="noprint">
	<p>
		<a href="#" onClick="fnPrint('print_pg');"><img src="/image/printer.jpg" width="39" height="36" border="0" alt="출력하기" /></a>
	</p>
</div>
<!--<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>-->
<div id="print_pg">

<table width="690" cellpadding="0" cellspacing="0">
  <tr>
    <td class="style32BC">인사기록카드</td>
  </tr>
  <tr>
    <td height="20" class="style20C">&nbsp;</td>
  </tr>
</table>
<table width="690" border="1px" cellpadding="15" cellspacing="0" bordercolor="#000000">
  <tr>
    <td style="border-bottom:none; border-top:none;">
     <table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td rowspan="4" class="first">
			<img src="<%=photo_image%>" width="110" height="120" alt="">
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">사원번호</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_no")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">성명(한글)</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_name")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">성명(영문)</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_ename")%></strong>
		</td>
      </tr>
      <tr>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">소속</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_org_code")%>)&nbsp;<%=rs("org_bonbu")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">직급(위)</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;(<%=rs("emp_grade")%>)&nbsp;<%=rs("emp_job")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">직책</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_position")%></strong>
		</td>
      </tr>
      <tr>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">입사일</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_in_date")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">주민번호</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_person1")%>-<%=rs("emp_person2")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">생년월일</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_birthday")%>&nbsp;&nbsp;(<%=rs("emp_sex")%>)</strong>
		</td>
      </tr>
      <tr>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">전화번호</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_tel_ddd")%>-<%=rs("emp_tel_no1")%>-<%=rs("emp_tel_no2")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">휴대폰번호</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">이메일주소</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=emp_email%></strong>
		</td>
      </tr>
      <tr>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">주소(현)</span>
		</td>
        <td width="62%" height="30" colspan="4" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%></strong>
		</td>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">비상연락)</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_emergency_tel")%></strong>
		</td>
      </tr>
      <tr>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">본적</span>
		</td>
        <td width="91%" height="30" colspan="6" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_family_sido")%>&nbsp;<%=rs("emp_family_gugun")%>&nbsp;<%=rs("emp_family_dong")%>&nbsp;<%=rs("emp_family_addr")%></strong>
		</td>
     </tr>
    </table>
   </td>
  </tr>
  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 학력사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;"><table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="22%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">기간</span>
		</td>
        <td width="18%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">학교명</span>
		</td>
        <td width="18%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">학과</span>
		</td>
        <td width="17%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">전공</span>
		</td>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">부전공</span>
		</td>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">학위</span>
		</td>
        <td width="7%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">졸업</span>
		</td>
      </tr>
   <% For i = 1 To k_sch %>
      <tr>
        <td width="22%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i, 1)%>&nbsp;~&nbsp;<%=sch_tab(i, 2)%></strong>
		</td>
        <td width="18%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i,3)%></strong>&nbsp;
		</td>
        <td width="18%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i,4)%></strong>&nbsp;
		</td>
        <td width="17%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i,5)%></strong>&nbsp;
		</td>
        <td width="9%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i,6)%></strong>&nbsp;
		</td>
        <td width="9%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i,7)%></strong>&nbsp;
		</td>
        <td width="7%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sch_tab(i,8)%></strong>&nbsp;
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>
  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 경력사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;"><table width="680" border="1px" cellpadding="0" cellspacing="0" >
      <tr>
        <td width="22%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">재직기간</span>
		</td>
        <td width="18%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">회사명</span>
		</td>
        <td width="18%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">부  서</span>
		</td>
        <td width="12%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">직위</span>
		</td>
        <td width="30%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">담당업무</span>
		</td>
      </tr>
   <% For i = 1 To k_car	%>
      <tr>
        <td width="22%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=car_tab(i,1)%>&nbsp;~&nbsp;<%=car_tab(i,2)%></strong>
		</td>
        <td width="18%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=car_tab(i,3)%></strong>
		</td>
        <td width="18%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=car_tab(i,4)%></strong>
		</td>
        <td width="12%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=car_tab(i,5)%></strong>
		</td>
        <td width="30%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=car_tab(i,6)%></strong>
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>

  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 자격사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;"><table width="680" border="1px" cellpadding="0" cellspacing="0" >
      <tr>
        <td width="24%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">자격종목</span>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">등급</span>
		</td>
        <td width="15%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">합격일자</span>
		</td>
        <td width="26%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">발급기관</span>
		</td>
        <td width="25%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">자격등록번호</span>
		</td>
      </tr>
   <% For i = 1 To k_qul	%>
      <tr>
        <td width="24%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=qul_tab(i,1)%></strong>
		</td>
        <td width="10%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=qul_tab(i,2)%></strong>
		</td>
        <td width="15%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=qul_tab(i,3)%></strong>
		</td>
        <td width="26%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=qul_tab(i,4)%></strong>
		</td>
        <td width="25%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=qul_tab(i,5)%></strong>
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>
</table>
<p style='page-break-before:always'><br style='height:0; line-height:0'>
<table width="690" cellpadding="0" cellspacing="0">
</table>
<table width="690" border="1px" cellpadding="15" cellspacing="0" bordercolor="#000000">
  <tr>
    <td style="border-bottom:none; border-top:none;">
	<table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">병역복무기간</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=mid(emp_military_date1,1,7)%>~<%=mid(emp_military_date2,1,7)%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">병역유형</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_military_id")%> - <%=rs("emp_military_grade")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">면제사유</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_military_comm")%></strong>
		</td>
      </tr>
      <tr>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">결혼기념일</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=emp_marry_date%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">종교</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_faith")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">취미</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_hobby")%></strong>
		</td>
      </tr>
      <tr>
      <%
		If rs("emp_sawo_id") = "Y" Then
		   sawo_id = "가입"
		Else
		   sawo_id = "안함"
		End If
	  %>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">경조회</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=sawo_id%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">장애유형</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_disabled")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">장애등급</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_disab_grade")%></strong>
		</td>
      </tr>
      <tr>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">최초입사일</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_first_date")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">근속기산일</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_gunsok_date")%></strong>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">퇴직기산일</span>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=rs("emp_end_gisan")%></strong>
		</td>
      </tr>
    </table>
	</td>
  </tr>
  <%
  rs.Close : Set rs = Nothing
  DBConn.Close : Set DBConn = Nothing
  %>
  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 가족사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;">
	<table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">관계</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">성명</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">생년월일</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">직업</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">전화번호</span>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">동거여부</span>
		</td>
      </tr>
   <% For i = 1 To k_fam %>
      <tr>
        <td width="10%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=fam_tab(i,1)%></strong>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=fam_tab(i,2)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=fam_tab(i,3)%>(<%=fam_tab(i,4)%>)</strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=fam_tab(i,5)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=fam_tab(i,6)%></strong>&nbsp;
		</td>
        <td width="10%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=fam_tab(i,7)%></strong>&nbsp;
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>
  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 어학사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;">
	<table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">어학구분</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">어학종류</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">점수</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">급수</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">취득일</span>
		</td>
        <td width="10%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">비  고</span>
		</td>
      </tr>
   <% For i = 1 To k_lan %>
      <tr>
        <td width="10%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=lan_tab(i,1)%></strong>
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=lan_tab(i,2)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=lan_tab(i,3)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=lan_tab(i,4)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=lan_tab(i,5)%></strong>&nbsp;
		</td>
        <td width="10%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;</strong>&nbsp;
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>
  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 교육사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;">
	<table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">교육과정명</span>
		</td>
        <td width="15%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">교육기관</span>
		</td>
        <td width="15%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">수료증번호</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">교육기간</span>
		</td>
        <td width="30%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">교육주요내용</span>
		</td>
      </tr>
   <% For i = 1 To k_edu%>
      <tr>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=edu_tab(i,1)%></strong>
		</td>
        <td width="15%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=edu_tab(i,2)%></strong>&nbsp;
		</td>
        <td width="15%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=edu_tab(i,3)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=edu_tab(i,4)%> - <%=edu_tab(i,5)%></strong>&nbsp;
		</td>
        <td width="30%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style1"><strong>&nbsp;<%=edu_tab(i,6)%></strong>&nbsp;
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>
  <tr>
    <td class="style14L" style="border-bottom:none; border-top:none;">❐ 발령사항</td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;"><table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="8%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">발령일</span>
		</td>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">발령구분</span>
		</td>
        <td width="9%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">발령유형</span>
		</td>
        <td width="14%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">회사</span>
		</td>
        <td width="11%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">소속</span>
		</td>
        <td width="12%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">직급/책</span>
		</td>
        <td width="16%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">기간</span>
		</td>
        <td width="20%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;">
			<span class="style1">비고</span>
		</td>
      </tr>
   <% For i = 1 To k_app %>
      <tr>
        <td width="8%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,1)%></strong>
		</td>
        <td width="9%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,2)%></strong>&nbsp;
		</td>
        <td width="9%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,3)%></strong>&nbsp;
		</td>
        <td width="14%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,11)%></strong>&nbsp;
		</td>
        <td width="11%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,12)%>)<%=app_tab(i,13)%></strong>&nbsp;
		</td>
        <td width="12%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,14)%> - <%=app_tab(i,16)%></strong>&nbsp;
		</td>
        <td width="16%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,18)%>-<%=app_tab(i,19)%>&nbsp;<%=app_tab(i,17)%></strong>&nbsp;
		</td>
        <td width="20%" height="30" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
			<span class="style2"><strong>&nbsp;<%=app_tab(i,20)%>&nbsp;<%=app_tab(i,21)%></strong>&nbsp;
		</td>
      </tr>
	<%	Next	%>
    </table>
	</td>
  </tr>
</table>

</div>

</p>
</body>
</html>