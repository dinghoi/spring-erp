<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next
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
Dim m_seq, emp_name, emp_ename, emp_birthday, emp_birthday_id, emp_org_code, emp_org_name
Dim emp_bonbu, emp_saupbu, emp_team, emp_reside_place, emp_org_level
Dim emp_type, emp_grade, emp_job, emp_position, emp_jikmu, emp_first_date, emp_in_date
Dim emp_gunsok_date, emp_yuncha_date, emp_end_gisan, emp_person1, emp_person2
Dim emp_sex, emp_tel_ddd, emp_tel_no1, emp_tel_no2, emp_hp_ddd, emp_hp_no1, emp_hp_no2
Dim emp_sido, emp_gugun, emp_dong, emp_addr, emp_zipcode, emp_email, emp_sawo_id, emp_sawo_date
Dim emp_marry_date, emp_hobby, emp_disabled, emp_disab_grade, emp_military_id, emp_military_grade
Dim emp_military_date1, emp_military_date2, emp_military_comm, emp_faith, emp_stay_name, emp_stay_code
Dim cost_group, emp_emergency_tel, emp_extension_no, emp_last_edu, cost_center
Dim emp_reside_company, emp_pay_id, emp_image, dz_id, emp_org_baldate, emp_disabled_yn
Dim emp_pay_type, emp_nation_code, kone_email, emp_hp, emp_grade_date, emp_end_date
Dim rsEmp, stock_end_date, stock_level, end_msg, rsDz

emp_no = f_toString(Request.Form("emp_no"), "")

'사번 검증 조회
objBuilder.Append "SELECT emp_no FROM emp_master WHERE emp_no = '"&emp_no&"';"

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsEmp.EOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('중복된 사번이 존재합니다.\n확인 후 다시 등록해주세요.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If

rsEmp.Close() : Set rsEmp = Nothing

emp_company = Request.Form("emp_company")
dz_id = f_toString(Request.Form("dz_id"), "")

'급여ID(더존) 검증 조회
objBuilder.Append "SELECT emtt.emp_no FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN dz_pay_info AS dpit ON emtt.emp_no = dpit.emp_no "
objBuilder.Append "WHERE emtt.emp_company = '"&emp_company&"' AND dpit.dz_id ='"&dz_id&"';"

Set rsDz = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsDz.EOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('중복된 급여ID 혹은 이미 등록된 사번이 존재합니다.\n확인 후 다시 등록해주세요.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If

rsDz.Close() : Set rsDz = Nothing

m_seq = Request.Form("m_seq")

emp_name = Request.Form("emp_name")
emp_ename = Request.Form("emp_ename")
emp_birthday = Request.Form("emp_birthday")
emp_birthday_id = Request.Form("emp_birthday_id")
emp_org_code = Request.Form("emp_org_code")
emp_org_name = Request.Form("emp_org_name")

emp_bonbu = Request.Form("emp_bonbu")
emp_saupbu = Request.Form("emp_saupbu")
emp_team = Request.Form("emp_team")
emp_reside_place = Request.Form("emp_reside_place")
emp_org_level = Request.Form("emp_org_level")
emp_type = Request.Form("emp_type")
emp_grade = Request.Form("emp_grade")
emp_job = Request.Form("emp_job")
emp_position = Request.Form("emp_position")
emp_jikmu = Request.Form("emp_jikmu")
emp_first_date = Request.Form("emp_first_date")
emp_in_date = Request.Form("emp_in_date")
emp_gunsok_date = Request.Form("emp_gunsok_date")
emp_yuncha_date = Request.Form("emp_yuncha_date")
emp_end_gisan = Request.Form("emp_end_gisan")
emp_person1 = Request.Form("emp_person1")
emp_person2 = Request.Form("emp_person2")
emp_sex = Request.Form("emp_sex")
emp_tel_ddd = Request.Form("emp_tel_ddd")
emp_tel_no1 = Request.Form("emp_tel_no1")
emp_tel_no2 = Request.Form("emp_tel_no2")
emp_hp_ddd = Request.Form("emp_hp_ddd")
emp_hp_no1 = Request.Form("emp_hp_no1")
emp_hp_no2 = Request.Form("emp_hp_no2")
emp_sido = Request.Form("emp_sido")
emp_gugun = Request.Form("emp_gugun")
emp_dong = Request.Form("emp_dong")
emp_addr = Request.Form("emp_addr")
emp_zipcode = Request.Form("emp_zipcode")
emp_email = Request.Form("emp_email")
emp_sawo_id = Request.Form("emp_sawo_id")
emp_sawo_date = Request.Form("emp_sawo_date")
emp_marry_date = Request.Form("emp_marry_date")
emp_hobby = Request.Form("emp_hobby")
emp_disabled = Request.Form("emp_disabled")
emp_disab_grade = Request.Form("emp_disab_grade")
emp_military_id = Request.Form("emp_military_id")
emp_military_grade = Request.Form("emp_military_grade")
emp_military_date1 = Request.Form("emp_military_date1")
emp_military_date2 = Request.Form("emp_military_date2")
emp_military_comm = Request.Form("emp_military_comm")
emp_faith = Request.Form("emp_faith")
emp_stay_name = Request.Form("emp_stay_name")
emp_stay_code = Request.Form("emp_stay_code")
cost_group = Request.Form("cost_group")
emp_emergency_tel = Request.Form("emp_emergency_tel")
emp_extension_no = Request.Form("emp_extension_no")
emp_last_edu = Request.Form("emp_last_edu")
cost_center = Request.Form("cost_center")
mg_group = Request.Form("mg_group")
emp_reside_company = Request.Form("emp_reside_company")
emp_pay_id = Request.Form("emp_pay_id")
emp_image = Request.Form("emp_image")

emp_org_baldate = Request.Form("emp_org_baldate")
emp_end_date = Request.Form("emp_end_date")
emp_grade_date = Request.Form("emp_grade_date")

If emp_org_level = "상주처" Then
	reside = "1"
	cost_center = "상주직접비"
Else
	reside = "0"
End If

If emp_disabled = "해당사항없음" Or emp_disabled = "" Then
	emp_disabled_yn = "N"
	emp_disab_grade = ""
Else
	emp_disabled_yn = "Y"
End If

If emp_sawo_id = "Y" Then
	emp_sawo_date = emp_in_date
Else
	emp_sawo_date = "1900-01-01"
End If

If cost_center = "상주직접비" Then
	If f_toString(cost_group, "") = "" Then
		cost_group =  emp_reside_company
	End If
End If

emp_pay_type = "1"
emp_nation_code = "001"

kone_email = emp_email&"@k-one.co.kr"
emp_hp = emp_hp_ddd&"-"&emp_hp_no1&"-"&emp_hp_no2

if f_toString(emp_sawo_date, "")= "" Then
   emp_sawo_date = "1900-01-01"
End If

If f_toString(emp_birthday, "") = "" Then
   emp_birthday = "1900-01-01"
End If

If f_toString(emp_end_date, "") = "" Then
   emp_end_date = "1900-01-01"
End If

If f_toString(emp_org_baldate, "") = "" Then
   emp_org_baldate = "1900-01-01"
End If

If f_toString(emp_grade_date, "") = "" Then
   emp_grade_date = "1900-01-01"
End If

If f_toString(emp_military_date1, "") = "" Then
   emp_military_date1 = "1900-01-01"
End If

If f_toString(emp_military_date2, "") = "" Then
   emp_military_date2 = "1900-01-01"
End If

If f_toString(emp_marry_date, "") = "" Then
   emp_marry_date = "1900-01-01"
End If

DBConn.BeginTrans

'직원 정보 등록
objBuilder.Append "INSERT INTO emp_master(emp_no,emp_name,emp_ename,emp_type,emp_sex,emp_person1,emp_person2,emp_image,emp_first_date,"
objBuilder.Append "emp_in_date,emp_gunsok_date,emp_yuncha_date,emp_end_gisan,emp_company,emp_bonbu,emp_saupbu,emp_team,"
objBuilder.Append "emp_org_code,emp_org_name,emp_stay_code,emp_stay_name,emp_reside_place,emp_reside_company,emp_grade,"
objBuilder.Append "emp_job,emp_position,emp_jikmu,emp_birthday,emp_birthday_id,"
objBuilder.Append "emp_zipcode,emp_sido,emp_gugun,emp_dong,emp_addr,"
objBuilder.Append "emp_tel_ddd,emp_tel_no1,emp_tel_no2,emp_hp_ddd,emp_hp_no1,emp_hp_no2,emp_email,emp_military_id,"
objBuilder.Append "emp_military_date1,emp_military_date2,emp_military_grade,emp_military_comm,emp_hobby,emp_faith,emp_last_edu,"
objBuilder.Append "emp_marry_date,emp_disabled_yn,emp_disabled,emp_disab_grade,emp_sawo_id,emp_sawo_date,emp_emergency_tel,"
objBuilder.Append "emp_extension_no,emp_nation_code,emp_pay_id,emp_pay_type,cost_center,cost_group,emp_reg_date,emp_reg_user)"
objBuilder.Append "VALUES('"&emp_no&"','"&emp_name&"','"&emp_ename&"','"&emp_type&"','"&emp_sex&"','"&emp_person1&"','"&emp_person2&"','"&emp_image&"','"&emp_first_date&"',"
objBuilder.Append "'"&emp_in_date&"','"&emp_gunsok_date&"','"&emp_yuncha_date&"','"&emp_end_gisan&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"',"
objBuilder.Append "'"&emp_org_code&"','"&emp_org_name&"','"&emp_stay_code&"','"&emp_stay_name&"','"&emp_reside_place&"','"&emp_reside_company&"','"&emp_grade&"',"
objBuilder.Append "'"&emp_job&"','"&emp_position&"','"&emp_jikmu&"','"&emp_birthday&"','"&emp_birthday_id&"',"
objBuilder.Append "'"&emp_zipcode&"','"&emp_sido&"','"&emp_gugun&"','"&emp_dong&"','"&emp_addr&"',"
objBuilder.Append "'"&emp_tel_ddd&"','"&emp_tel_no1&"','"&emp_tel_no2&"','"&emp_hp_ddd&"','"&emp_hp_no1&"','"&emp_hp_no2&"','"&emp_email&"','"&emp_military_id&"',"
objBuilder.Append "'"&emp_military_date1&"','"&emp_military_date2&"','"&emp_military_grade&"','"&emp_military_comm&"','"&emp_hobby&"','"&emp_faith&"','"&emp_last_edu&"',"
objBuilder.Append "'"&emp_marry_date&"','"&emp_disabled_yn&"','"&emp_disabled&"','"&emp_disab_grade&"','"&emp_sawo_id&"','"&emp_sawo_date&"','"&emp_emergency_tel&"',"
objBuilder.Append "'"&emp_extension_no&"','"&emp_nation_code&"','"&emp_pay_id&"','"&emp_pay_type&"','"&cost_center&"','"&cost_group&"',NOW(),'"&user_name&"');"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'더존 급여 정보 등록
objBuilder.Append "INSERT INTO dz_pay_info(dz_id, emp_company, emp_no , reg_id)"
objBuilder.Append "VALUES('"&dz_id&"', '"&emp_company&"', '"&emp_no&"', '"&user_id&"');"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'경조회 정보 등록
If emp_sawo_id = "Y" Then
	objBuilder.Append "INSERT INTO emp_sawo_mem(sawo_empno,sawo_date,sawo_id,sawo_emp_name,sawo_company,"
	objBuilder.Append "sawo_orgcode,sawo_org_name,sawo_target,sawo_in_count,sawo_in_pay,sawo_give_count,sawo_give_pay, sawo_reg_date,sawo_reg_user)"
	objBuilder.Append"VALUES('"&emp_no&"','"&emp_sawo_date&"','입사','"&emp_name&"','"&emp_company&"',"
	objBuilder.Append "'"&emp_org_code&"','"&emp_org_name&"','Y',0,0,0,0,NOW(),'"&user_id&"');"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

'회원 정보 등록
objBuilder.Append "INSERT INTO memb(user_id,pass,emp_no,user_name,user_grade,"
objBuilder.Append "position,emp_company,bonbu,saupbu,team,org_name,"
objBuilder.Append "hp,email,reside_place,reside_company,"
objBuilder.Append "reside,mg_group,grade,sms,help_yn,reg_date,reg_id,reg_name)"
objBuilder.Append "VALUES('"&emp_no&"','"&emp_person2&"','"&emp_no&"','"&emp_name&"','"&emp_job&"',"
objBuilder.Append "'"&emp_position&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"',"
objBuilder.Append "'"&emp_org_name&"','"&emp_hp&"','"&kone_email&"','"&emp_reside_place&"','"&emp_reside_company&"',"
objBuilder.Append "'"&reside&"','"&mg_group&"','4','N','N',NOW(),'"&user_id&"','"&user_name&"');"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'자산 정보 등록
stock_end_date = "1900-01-01"
stock_level = "개인"

objBuilder.Append "INSERT INTO met_stock_code(stock_code,stock_level,stock_name,stock_company,stock_bonbu,"
objBuilder.Append "stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code,"
objBuilder.Append "stock_manager_name,reg_date,reg_user)"
objBuilder.Append "VALUES('"&emp_no&"','"&stock_level&"','"&emp_name&"','"&emp_company&"','"&emp_bonbu&"',"
objBuilder.Append "'"&emp_saupbu&"','"&emp_team&"','"&emp_in_date&"','"&stock_end_date&"','"&emp_no&"',"
objBuilder.Append "'"&emp_name&"',NOW(),'"&user_name&"');"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'가족사항 등록
Dim rsFamily, arrFamily, i
Dim f_seq, f_rel, f_name, f_birthday, f_birthday_id, f_job, f_live
Dim f_person1, f_person2, f_tel_ddd, f_tel_no1, f_tel_no2, f_support_yn
Dim f_national, f_disab, f_merit, f_serius, f_pensioner, f_witak, f_holt
Dim f_holt_date, f_children

objBuilder.Append "SELECT f_seq, f_rel, f_name, f_birthday, f_birthday_id, f_job, f_live, f_person1, f_person2, "
objBuilder.Append "	f_tel_ddd, f_tel_no1, f_tel_no2, f_support_yn, f_national, f_disab, f_merit, f_serius, "
objBuilder.Append "	f_pensioner, f_witak, f_holt, f_holt_date, f_children "
objBuilder.Append "FROM member_family "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsFamily = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsFamily.EOF Then
	arrFamily = rsFamily.getRows()
End If
rsFamily.Close() : Set rsFamily = Nothing

If IsArray(arrFamily) Then
	For i = LBound(arrFamily) To UBound(arrFamily, 2)
		f_seq = arrFamily(0, i)
		f_rel = arrFamily(1, i)
		f_name = arrFamily(2, i)
		f_birthday = arrFamily(3, i)
		f_birthday_id = arrFamily(4, i)
		f_job = arrFamily(5, i)
		f_live = arrFamily(6, i)
		f_person1 = arrFamily(7, i)
		f_person2 = arrFamily(8, i)
		f_tel_ddd = arrFamily(9, i)
		f_tel_no1 = arrFamily(10, i)
		f_tel_no2 = arrFamily(11, i)
		f_support_yn = arrFamily(12, i)
		f_national = arrFamily(13, i)
		f_disab = arrFamily(14, i)
		f_merit = arrFamily(15, i)
		f_serius = arrFamily(16, i)
		f_pensioner = arrFamily(17, i)
		f_witak = arrFamily(18, i)
		f_holt = arrFamily(19, i)
		f_holt_date = arrFamily(20, i)
		f_children = arrFamily(21, i)

		objBuilder.Append "INSERT INTO emp_family(family_empno,family_seq,family_rel,family_name,family_birthday,"
		objBuilder.Append "family_birthday_id,family_job,family_live,family_support_yn,family_person1,"
		objBuilder.Append "family_person2,family_tel_ddd,family_tel_no1,family_tel_no2,family_national,"
		objBuilder.Append "family_disab,family_merit,family_serius,family_pensioner,family_witak,"
		objBuilder.Append "family_holt,family_holt_date,family_children,family_reg_date,family_reg_user)"
		objBuilder.Append "VALUES('"&emp_no&"','"&f_seq&"','"&f_rel&"','"&f_name&"','"&f_birthday&"',"
		objBuilder.Append "'"&f_birthday_id&"','"&f_job&"','"&f_live&"','"&f_support_yn&"','"&f_person1&"',"
		objBuilder.Append "'"&f_person2&"','"&f_tel_ddd&"','"&f_tel_no1&"','"&f_tel_no2&"','"&f_national&"',"
		objBuilder.Append "'"&f_disab&"','"&f_merit&"','"&f_serius&"','"&f_pensioner&"','"&f_witak&"',"
		objBuilder.Append "'"&f_holt&"','"&f_holt_date&"','"&f_children&"',NOW(),'"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'학력사항 등록
Dim rsSch, arrSch, sch_seq, sch_start_date, sch_end_date, sch_school_name, sch_dept
Dim sch_major, sch_sub_major, sch_degree, sch_finish, sch_comment

objBuilder.Append "SELECT sch_seq, sch_start_date, sch_end_date, sch_school_name, sch_dept, "
objBuilder.Append "	sch_major, sch_sub_major, sch_degree, sch_finish, sch_comment "
objBuilder.Append "FROM member_school "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsSch = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSch.EOF Then
	arrSch = rsSch.getRows()
End If
rsSch.Close() : Set rsSch = Nothing

If IsArray(arrSch) Then
	For i = LBound(arrSch) To UBound(arrSch, 2)
		sch_seq = arrSch(0, i)
		sch_start_date = arrSch(1, i)
		sch_end_date = arrSch(2, i)
		sch_school_name = arrSch(3, i)
		sch_dept = arrSch(4, i)
		sch_major = arrSch(5, i)
		sch_sub_major = arrSch(6, i)
		sch_degree = arrSch(7, i)
		sch_finish = arrSch(8, i)
		sch_comment = arrSch(9, i)

		objBuilder.Append "INSERT INTO emp_school(sch_empno,sch_seq,sch_start_date,sch_end_date,"
		objBuilder.Append "sch_school_name,sch_dept,sch_major,sch_sub_major,sch_degree,"
		objBuilder.Append "sch_finish,sch_comment,sch_reg_date,sch_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&sch_seq&"','"&sch_start_date&"','"&sch_end_date&"','"&sch_school_name&"',"
		objBuilder.Append "'"&sch_dept&"','"&sch_major&"','"&sch_sub_major&"','"&sch_degree&"','"&sch_finish&"',"
		objBuilder.Append "'"&sch_comment&"',NOW(),'"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'경력사항 등록
Dim rsCareer, arrCareer, c_seq, c_join_date, c_end_date, c_office, c_dept
Dim c_position, c_task

objBuilder.Append "SELECT c_seq, c_join_date, c_end_date, c_office, c_dept, c_position, c_task "
objBuilder.Append "FROM member_career "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsCareer = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCareer.EOF Then
	arrCareer = rsCareer.getRows()
End If
rsCareer.Close() : Set rsCareer = Nothing

If IsArray(arrCareer) Then
	For i = LBound(arrCareer) To UBound(arrCareer, 2)
		c_seq = arrCareer(0, i)
		c_join_date = arrCareer(1, i)
		c_end_date = arrCareer(2, i)
		c_office = arrCareer(3, i)
		c_dept = arrCareer(4, i)
		c_position = arrCareer(5, i)
		c_task = arrCareer(6, i)

		objBuilder.Append "INSERT INTO emp_career(career_empno,career_seq,career_join_date,career_end_date,career_office,"
		objBuilder.Append "career_dept,career_position,career_task,career_reg_date,career_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&c_seq&"','"&c_join_date&"','"&c_end_date&"','"&c_office&"',"
		objBuilder.Append "'"&c_dept&"','"&c_position&"','"&c_task&"',NOW(),'"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'자격사항 등록
Dim rsQual, arrQual, qual_seq, qual_type, qual_grade, qual_pass_date
Dim qual_org, qual_no, qual_passport, qual_pay_id

objBuilder.Append "SELECT qual_seq, qual_type, qual_grade, qual_pass_date, qual_org, qual_no, qual_passport, qual_pay_id "
objBuilder.Append "FROM member_qual "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsQual = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsQual.EOF Then
	arrQual = rsQual.getRows()
End If
rsQual.Close() : Set rsQual = Nothing

If IsArray(arrQual) Then
	For i = LBound(arrQual) To UBound(arrQual, 2)
		qual_seq = arrQual(0, i)
		qual_type = arrQual(1, i)
		qual_grade = arrQual(2, i)
		qual_pass_date = arrQual(3, i)
		qual_org = arrQual(4, i)
		qual_no = arrQual(5, i)
		qual_passport = arrQual(6, i)
		qual_pay_id = arrQual(7, i)

		objBuilder.Append "INSERT INTO emp_qual(qual_empno,qual_seq,qual_type,qual_grade,qual_pass_date, "
		objBuilder.Append "qual_org,qual_no,qual_passport,qual_pay_id,qual_reg_date,qual_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&qual_seq&"','"&qual_type&"','"&qual_grade&"','"&qual_pass_date&"',"
		objBuilder.Append "'"&qual_org&"','"&qual_no&"','"&qual_passport&"','"&qual_pay_id&"',NOW(),'"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'교육사항 등록
Dim rsEdu, arrEdu, edu_seq, edu_name, edu_office, edu_finish_no
Dim edu_start_date, edu_end_date, edu_pay, edu_comment

objBuilder.Append "SELECT edu_seq, edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date, edu_pay, edu_comment "
objBuilder.Append "FROM member_edu "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsEdu = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsEdu.EOF Then
	arrEdu = rsEdu.getRows()
End If
rsEdu.Close() : Set rsEdu = Nothing

If IsArray(arrEdu) Then
	For i = LBound(arrEdu) To UBound(arrEdu, 2)
		edu_seq = arrEdu(0, i)
		edu_name = arrEdu(1, i)
		edu_office = arrEdu(2, i)
		edu_finish_no = arrEdu(3, i)
		edu_start_date = arrEdu(4, i)
		edu_end_date = arrEdu(5, i)
		edu_pay = arrEdu(6, i)
		edu_comment = arrEdu(7, i)

		objBuilder.Append "INSERT INTO emp_edu (edu_empno,edu_seq,edu_name,edu_office,edu_finish_no,"
		objBuilder.Append "edu_start_date,edu_end_date,edu_pay,edu_comment,edu_reg_date,edu_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"',"
		objBuilder.Append "'"&edu_start_date&"','"&edu_end_date&"','"&edu_pay&"','"&edu_comment&"',NOW(),'"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'어학능력 등록
Dim rsLang, arrLang, lang_seq, lang_id, lang_id_type, lang_point, lang_grade, lang_get_date

objBuilder.Append "SELECT lang_seq, lang_id, lang_id_type, lang_point, lang_grade, lang_get_date "
objBuilder.Append "FROM member_language "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsLang = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsLang.EOF Then
	arrLang = rsLang.getRows()
End If
rsLang.Close() : Set rsLang = Nothing

If IsArray(arrLang) Then
	For i = LBound(arrLang) To UBound(arrLang, 2)
		lang_seq = arrLang(0, i)
		lang_id = arrLang(1, i)
		lang_id_type = arrLang(2, i)
		lang_point = arrLang(3, i)
		lang_grade = arrLang(4, i)
		lang_get_date = arrLang(5, i)

		objBuilder.Append "INSERT INTO emp_language(lang_empno,lang_seq,lang_id,lang_id_type,lang_point,"
		objBuilder.Append "lang_grade,lang_get_date,lang_reg_date,lang_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&lang_seq&"','"&lang_id&"','"&lang_id_type&"','"&lang_point&"',"
		objBuilder.Append "'"&lang_grade&"','"&lang_get_date&"',NOW(),'"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'승인 완료 처리
Dim rsApprove

objBuilder.Append "SELECT emp_no FROM emp_master WHERE emp_no = '"&emp_no&"';"

Set rsApprove = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsApprove.EOF Or rsApprove.BOF Then
	DBConn.RollbackTrans

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('처리 진행 중 비정상적인 오류가 발생했습니다.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
Else
	objBuilder.Append "UPDATE member_info SET m_approve_yn = 'Y' WHERE m_seq = '"&m_seq&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If
rsApprove.Close() : Set rsApprove = Nothing

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리 중 Error가 발생했습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 승인 처리되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>