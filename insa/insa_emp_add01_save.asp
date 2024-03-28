<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->

<!--include file="../xmlrpc.asp"-->
<!--include file="../class.EmmaSMS.asp"-->
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
Dim uploadForm, u_type, emp_name, emp_ename, emp_type, dz_id
Dim emp_sex, emp_person1, emp_person2, sex_id, emp_first_date
Dim emp_in_date, emp_gunsok_date, emp_yuncha_date, emp_end_gisan
Dim emp_end_date, emp_bonbu, emp_saupbu, emp_team
Dim emp_org_code, emp_org_name, emp_org_baldate, emp_stay_code, emp_stay_name
Dim emp_reside_place, emp_reside_company, emp_org_level
Dim emp_grade, emp_grade_date, emp_job, emp_position, emp_jikmu
Dim emp_birthday, emp_birthday_id
Dim emp_zipcode, emp_sido, emp_gugun, emp_dong, emp_addr
Dim emp_tel_ddd, emp_tel_no1, emp_tel_no2, emp_hp_ddd, emp_hp_no1, emp_hp_no2
Dim emp_email, emp_military_id, emp_military_date1, emp_military_date2, emp_military_grade
Dim emp_military_comm, emp_hobby, emp_faith, emp_marry_date, emp_disabled, emp_disab_grade
Dim emp_disabled_yn, emp_sawo_id, emp_sawo_date
Dim emp_emergency_tel, emp_extension_no, emp_last_edu, cost_center, cost_group
Dim emp_pay_id, emp_pay_type, emp_nation_code, emp_hp, kone_email
Dim rs_sawo, rs_memb, rs_stock, rsDz
Dim stock_end_date, stock_level, end_msg
Dim emp_jikgun, rsEmp

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")

uploadForm.AbsolutePath = True
uploadForm.Overwrite = True
uploadForm.MaxUploadSize = 1024*1024*50

u_type = uploadForm("u_type")
emp_no = uploadForm("emp_no")
dz_id = uploadForm("dz_id")
emp_company = uploadForm("emp_company")

If u_type <> "U" Then
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
End If

'급여ID(더존) 검증 조회
objBuilder.Append "SELECT emtt.emp_no FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN dz_pay_info AS dpit ON emtt.emp_no = dpit.emp_no "
objBuilder.Append "WHERE emtt.emp_company = '"&emp_company&"' AND dpit.dz_id ='"&dz_id&"' "

If u_type = "U" Then
	objBuilder.Append "	AND dpit.emp_no <> '"&emp_no&"';"
End If

Set rsDz = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsDz.EOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('중복된 급여ID가 존재합니다.\n확인 후 다시 등록해주세요.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If
rsDz.Close() : Set rsDz = Nothing


emp_name = uploadForm("emp_name")
emp_ename = uploadForm("emp_ename")
emp_type = uploadForm("emp_type")
emp_sex = uploadForm("emp_sex")
emp_person1 = uploadForm("emp_person1")
emp_person2 = uploadForm("emp_person2")

If emp_person2 <> "" Then
   sex_id = Mid(CStr(emp_person2), 1, 1)

	If sex_id = "1" Then
		 emp_sex = "남"
	Else
		 emp_sex = "여"
	End If
End If

emp_first_date = uploadForm("emp_first_date")
emp_in_date = uploadForm("emp_in_date")
emp_gunsok_date = uploadForm("emp_gunsok_date")
emp_yuncha_date = uploadForm("emp_yuncha_date")
emp_end_gisan = uploadForm("emp_end_gisan")
emp_end_date = uploadForm("emp_end_date")
emp_bonbu = uploadForm("emp_bonbu")
emp_saupbu = uploadForm("emp_saupbu")
emp_team = uploadForm("emp_team")
emp_org_code = uploadForm("emp_org_code")
emp_org_name = uploadForm("emp_org_name")
emp_org_baldate = uploadForm("emp_org_baldate")
emp_stay_code = uploadForm("emp_stay_code")
emp_stay_name = uploadForm("emp_stay_name")
emp_reside_place = uploadForm("emp_reside_place")
emp_reside_company = uploadForm("emp_reside_company")
emp_org_level = uploadForm("emp_org_level")

If emp_org_level = "상주처" Then
	reside = "1"
Else
	reside = "0"
End If

'Dim emp_family_zip, emp_family_sido, emp_family_gugun, emp_family_dong, emp_family_addr
'emp_family_zip = uploadForm("emp_family_zip")
'emp_family_sido = uploadForm("emp_family_sido")
'emp_family_gugun = uploadForm("emp_family_gugun")
'emp_family_dong = uploadForm("emp_family_dong")
'emp_family_addr = uploadForm("emp_family_addr")

emp_grade = uploadForm("emp_grade")
emp_grade_date = uploadForm("emp_grade_date")
emp_job = uploadForm("emp_job")
emp_position = uploadForm("emp_position")
emp_jikmu = uploadForm("emp_jikmu")
emp_birthday = uploadForm("emp_birthday")
emp_birthday_id = uploadForm("emp_birthday_id")
emp_zipcode = uploadForm("emp_zipcode")
emp_sido = uploadForm("emp_sido")
emp_gugun = uploadForm("emp_gugun")
emp_dong = uploadForm("emp_dong")
emp_addr = uploadForm("emp_addr")
emp_tel_ddd = uploadForm("emp_tel_ddd")
emp_tel_no1 = uploadForm("emp_tel_no1")
emp_tel_no2 = uploadForm("emp_tel_no2")
emp_hp_ddd = uploadForm("emp_hp_ddd")
emp_hp_no1 = uploadForm("emp_hp_no1")
emp_hp_no2 = uploadForm("emp_hp_no2")
emp_email = uploadForm("emp_email")
emp_military_id = uploadForm("emp_military_id")
emp_military_date1 = uploadForm("emp_military_date1")
emp_military_date2 = uploadForm("emp_military_date2")
emp_military_grade = uploadForm("emp_military_grade")
emp_military_comm = uploadForm("emp_military_comm")
emp_hobby = uploadForm("emp_hobby")
emp_faith = uploadForm("emp_faith")
emp_marry_date = uploadForm("emp_marry_date")
emp_disabled = uploadForm("emp_disabled")
emp_disab_grade = uploadForm("emp_disab_grade")

If emp_disabled = "해당사항없음" Or emp_disabled = "" Then
	emp_disabled_yn = "N"
	emp_disab_grade = ""
Else
	emp_disabled_yn = "Y"
End If

emp_sawo_id = uploadForm("emp_sawo_id")

If emp_sawo_id = "Y" Then
	If u_type = "U" Then
		emp_sawo_date = uploadForm("emp_sawo_date")
	Else
		emp_sawo_date = uploadForm("emp_in_date")
	End If
Else
	emp_sawo_date = "1900-01-01"
End If

emp_emergency_tel = uploadForm("emp_emergency_tel")
emp_extension_no = uploadForm("emp_extension_no")
emp_last_edu = uploadForm("emp_last_edu")
cost_center = uploadForm("cost_center")
cost_group = uploadForm("cost_group")

If emp_org_level = "상주처" Then
	cost_center = "상주직접비"
End If

If cost_center = "상주직접비" Then
   If f_toString(cost_group, "") = "" Then
		cost_group =  emp_reside_company
   End If
End If

mg_group = uploadForm("mg_group")
emp_pay_id = uploadForm("emp_pay_id")

'	emp_pay_id = "0"
emp_pay_type = "1"
emp_nation_code = "001"
kone_email = emp_email&"@k-won.co.kr"
emp_hp = emp_hp_ddd&"-"&emp_hp_no1&"-"&emp_hp_no2

Dim v_att_file, path, filenm, filename, fileType, save_path

v_att_file= uploadForm("v_att_file")
path = Server.MapPath ("/emp_photo")

Set filenm = uploadForm("att_file")(1)
filename = filenm

If filenm.FileExists Then
	If filenm.length > 1024*1024*8  Then
		Response.Write "<script type='text/javascript'>"
		Response.Write "	alert('업로드 파일 최대 용량은 2MB를 넘을 수 없습니다.');"
		Response.Write "	history.go(-1);"
		Response.Write "</script>"
		Response.End
	Else
		filename = filenm.safeFileName
		fileType = Mid(filename, inStrRev(filename,".") + 1)
		filename = emp_name&"_"&emp_no&"photo."&fileType
		save_path = path & "\" & filename

		filenm.save save_path
	End If
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

If f_toString(emp_sawo_date, "") = "" Then
   emp_sawo_date = "1900-01-01"
End If

DBConn.BeginTrans

If u_type = "U" then
	objBuilder.Append "UPDATE emp_master SET "
	objBuilder.Append "	emp_name ='"&emp_name&"',emp_ename ='"&emp_ename&"',emp_type ='"&emp_type&"',emp_sex ='"&emp_sex&"',"
	objBuilder.Append "	emp_person1 ='"&emp_person1&"',emp_person2 ='"&emp_person2&"',emp_first_date ='"&emp_first_date&"',emp_in_date ='"&emp_in_date&"',"
	objBuilder.Append "	emp_gunsok_date ='"&emp_gunsok_date&"',emp_yuncha_date ='"&emp_yuncha_date&"',emp_end_gisan ='"&emp_end_gisan&"',emp_company ='"&emp_company&"',"
	objBuilder.Append "	emp_bonbu ='"&emp_bonbu&"',emp_saupbu ='"&emp_saupbu&"',emp_team ='"&emp_team&"',emp_org_code ='"&emp_org_code&"',"
	objBuilder.Append "	emp_org_name ='"&emp_org_name&"',emp_grade ='"&emp_grade&"',emp_job ='"&emp_job&"',emp_position ='"&emp_position&"',"
	objBuilder.Append "	emp_stay_code ='"&emp_stay_code&"',emp_stay_name ='"&emp_stay_name&"',emp_reside_place ='"&emp_reside_place&"',emp_reside_company ='"&emp_reside_company&"',"
	objBuilder.Append "	emp_jikmu ='"&emp_jikmu&"',emp_birthday ='"&emp_birthday&"',emp_birthday_id ='"&emp_birthday_id&"',"
	objBuilder.Append "	emp_zipcode ='"&emp_zipcode&"',emp_sido ='"&emp_sido&"',emp_gugun ='"&emp_gugun&"',emp_dong ='"&emp_dong&"',"
	objBuilder.Append "	emp_addr ='"&emp_addr&"',emp_tel_ddd ='"&emp_tel_ddd&"',emp_tel_no1 ='"&emp_tel_no1&"',emp_tel_no2 ='"&emp_tel_no2&"',"
	objBuilder.Append "	emp_hp_ddd ='"&emp_hp_ddd&"',emp_hp_no1 ='"&emp_hp_no1&"',emp_hp_no2 ='"&emp_hp_no2&"', emp_email ='"&emp_email&"',"
	objBuilder.Append "	emp_military_id ='"&emp_military_id&"',emp_military_date1 ='"&emp_military_date1&"', emp_military_date2 ='"&emp_military_date2&"',"
	objBuilder.Append "	emp_military_grade ='"&emp_military_grade&"',emp_military_comm ='"&emp_military_comm&"',emp_hobby ='"&emp_hobby&"',emp_faith ='"&emp_faith&"',"
	objBuilder.Append "	emp_last_edu ='"&emp_last_edu&"',emp_marry_date ='"&emp_marry_date&"',emp_disabled_yn ='"&emp_disabled_yn&"',emp_disabled ='"&emp_disabled&"',"
	objBuilder.Append "	emp_disab_grade ='"&emp_disab_grade&"',emp_sawo_id ='"&emp_sawo_id&"',emp_sawo_date ='"&emp_sawo_date&"', emp_pay_id='"&emp_pay_id&"', "
	objBuilder.Append "	emp_emergency_tel ='"&emp_emergency_tel&"',emp_extension_no ='"&emp_extension_no&"',cost_center ='"&cost_center&"',cost_group ='"&cost_group&"',"
	objBuilder.Append "	emp_mod_user = '"&user_name&"',emp_mod_date = NOW() "

	If filenm <> "" Then
		objBuilder.Append ", emp_image ='"&filename&"' "
	End If

	objBuilder.Append "WHERE emp_no ='"&emp_no&"';"
Else
	objBuilder.Append "INSERT INTO emp_master(emp_no, emp_name, emp_ename, emp_type, emp_sex,"
	objBuilder.Append "emp_person1, emp_person2, emp_first_date, emp_in_date, emp_gunsok_date,"
	objBuilder.Append "emp_yuncha_date, emp_end_gisan, emp_company, emp_bonbu, emp_saupbu,"
	objBuilder.Append "emp_team, emp_org_code, emp_org_name, emp_stay_code, emp_stay_name,"
	objBuilder.Append "emp_reside_place, emp_reside_company, emp_grade, emp_job, emp_position,"
	objBuilder.Append "emp_jikgun, emp_jikmu, emp_birthday, emp_birthday_id, emp_zipcode, "
	objBuilder.Append "emp_sido, emp_gugun, emp_dong, emp_addr, emp_tel_ddd,"
	objBuilder.Append "emp_tel_no1, emp_tel_no2, emp_hp_ddd, emp_hp_no1,"
	objBuilder.Append "emp_hp_no2, emp_email, emp_military_id, emp_military_date1, emp_military_date2,"
	objBuilder.Append "emp_military_grade, emp_military_comm, emp_hobby, emp_faith, emp_last_edu,"
	objBuilder.Append "emp_marry_date, emp_disabled_yn, emp_disabled, emp_disab_grade, emp_sawo_id,"
	objBuilder.Append "emp_sawo_date, emp_emergency_tel, emp_extension_no, emp_nation_code, emp_pay_id,"
	objBuilder.Append "emp_pay_type, cost_center, cost_group, emp_reg_date, emp_reg_user "

	If filenm <> "" Then
		objBuilder.Append ", emp_image"
	End If

	objBuilder.Append ")VALUES("
	objBuilder.Append "'"&emp_no&"','"&emp_name&"','"&emp_ename&"','"&emp_type&"','"&emp_sex&"',"
	objBuilder.Append "'"&emp_person1&"','"&emp_person2&"','"&emp_first_date&"','"&emp_in_date&"','"&emp_gunsok_date&"',"
	objBuilder.Append "'"&emp_yuncha_date&"','"&emp_end_gisan&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"',"
	objBuilder.Append "'"&emp_team&"','"&emp_org_code&"','"&emp_org_name&"','"&emp_stay_code&"','"&emp_stay_name&"',"
	objBuilder.Append "'"&emp_reside_place&"','"&emp_reside_company&"','"&emp_grade&"','"&emp_job&"','"&emp_position&"',"
	objBuilder.Append "'"&emp_jikgun&"','"&emp_jikmu&"','"&emp_birthday&"','"&emp_birthday_id&"','"&emp_zipcode&"',"
	objBuilder.Append "'"&emp_sido&"','"&emp_gugun&"','"&emp_dong&"','"&emp_addr&"','"&emp_tel_ddd&"',"
	objBuilder.Append "'"&emp_tel_no1&"','"&emp_tel_no2&"','"&emp_hp_ddd&"','"&emp_hp_no1&"',"
	objBuilder.Append "'"&emp_hp_no2&"','"&emp_email&"','"&emp_military_id&"','"&emp_military_date1&"','"&emp_military_date2&"',"
	objBuilder.Append "'"&emp_military_grade&"','"&emp_military_comm&"','"&emp_hobby&"','"&emp_faith&"','"&emp_last_edu&"',"
	objBuilder.Append "'"&emp_marry_date&"','"&emp_disabled_yn&"','"&emp_disabled&"','"&emp_disab_grade&"','"&emp_sawo_id&"',"
	objBuilder.Append "'"&emp_sawo_date&"','"&emp_emergency_tel&"','"&emp_extension_no&"','"&emp_nation_code&"','"&emp_pay_id&"',"
	objBuilder.Append "'"&emp_pay_type&"','"&cost_center&"','"&cost_group&"',NOW(),'"&user_name&"'"

	If filenm <> "" Then
		objBuilder.Append ",'"&filename&"'"
	End If

	objBuilder.Append ");"
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Dim rsDzInfo
'더존 급여 정보 등록/수정
If f_toString(dz_id, "") <> "" Then
	objBuilder.Append "SELECT dz_id FROM dz_pay_info WHERE emp_no = '"&emp_no&"';"

	Set rsDzInfo = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsDzInfo.EOF Or rsDzInfo.BOF Then
		objBuilder.Append "INSERT INTO dz_pay_info(dz_id, emp_company, emp_no, reg_id)"
		objBuilder.Append "VALUES('"&dz_id&"', '"&emp_company&"', '"&emp_no&"', '"&user_id&"');"
	Else
		objBuilder.Append "UPDATE dz_pay_info SET "
		objBuilder.Append "	dz_id='"&dz_id&"', emp_company='"&emp_company&"', mod_date=NOW(), mod_id='"&user_id&"' "
		objBuilder.Append "WHERE emp_no = '"&emp_no&"';"
	End If
	rsDzInfo.Close() : Set rsDzInfo = Nothing

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

'회원 정보 등록
If f_toString(emp_no, "") <> "" Then
	objBuilder.Append "SELECT user_id FROM memb WHERE user_id = '"&emp_no&"';"

	Set rs_memb = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_memb.EOF Or rs_memb.BOF Then
	   objBuilder.Append "INSERT INTO memb(user_id, pass, emp_no, user_name, user_grade,"
	   objBuilder.Append "position, emp_company, bonbu, saupbu, team,"
	   objBuilder.Append "org_name, hp, email, reside_place, reside_company,"
	   objBuilder.Append "reside, mg_group, grade, sms, help_yn, reg_date, reg_id, reg_name)"
	   objBuilder.Append "VALUES('"&emp_no&"','"&emp_person2&"','"&emp_no&"','"&emp_name&"','"&emp_job&"',"
	   objBuilder.Append "'"&emp_position&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"',"
	   objBuilder.Append "'"&emp_org_name&"','"&emp_hp&"','"&kone_email&"','"&emp_reside_place&"','"&emp_reside_company&"',"
	   objBuilder.Append "'"&reside&"','"&mg_group&"','4','N','N',now(),'"&user_id&"','"&user_name&"');"
	Else
		objBuilder.Append "UPDATE memb SET "
		objBuilder.Append "	user_name='"&emp_name&"',user_grade='"&emp_job&"',position='"&emp_position&"',emp_company='"&emp_company&"',"
		objBuilder.Append "	bonbu='"&emp_bonbu&"',saupbu='"&emp_saupbu&"',team='"&emp_team&"',org_name='"&emp_org_name&"',"
		objBuilder.Append "	hp='"&emp_hp&"',email='"&kone_email&"',reside_place='"&emp_reside_place&"',reside_company='"&emp_reside_company&"',"
		objBuilder.Append "	reside='"&reside&"',mg_group='"&mg_group&"',mod_id='"&user_id&"',mod_date=NOW() "
		objBuilder.Append "WHERE user_id='"&emp_no&"';"
	End If
	DBConn.Execute(objBuilder.ToString())
	objBUilder.Clear()

	rs_memb.Close() : Set rs_memb = Nothing
End If

'경조회 정보 등록
If emp_sawo_id = "Y" Then
   objBuilder.Append "SELECT sawo_empno FROM emp_sawo_mem WHERE sawo_empno = '"&emp_no&"';"

   Set rs_sawo = DBConn.Execute(objBuilder.ToString())
   objBuilder.Clear()

   If rs_sawo.EOF Or rs_sawo.BOF Then
		objBuilder.Append "INSERT INTO emp_sawo_mem(sawo_empno, sawo_date, sawo_id, sawo_emp_name, sawo_company,"
		objBuilder.Append "sawo_orgcode, sawo_org_name, sawo_target, sawo_in_count, sawo_in_pay, sawo_give_count, sawo_give_pay)"
		objBuilder.Append "VALUES('"&emp_no&"','"&emp_sawo_date&"','입사','"&emp_name&"','"&emp_company&"',"
		objBuilder.Append "'"&emp_org_code&"','"&emp_org_name&"','Y',0,0,0,0);"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If

	rs_sawo.Close() : Set rs_sawo = Nothing
End If

'자산 정보 등록
If f_toString(emp_no, "") <> "" Then
	objBuilder.Append "SELECT stock_code FROM met_stock_code WHERE stock_code='"&emp_no&"';"

	Set rs_stock = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_stock.EOF Or rs_stock.BOF Then
	   stock_end_date = "1900-01-01"
	   stock_level = "개인"

	   objBuilder.Append "INSERT INTO met_stock_code (stock_code, stock_level, stock_name, stock_company, "
	   objBuilder.Append "stock_bonbu, stock_saupbu, stock_team, stock_open_date, "
	   objBuilder.Append "stock_end_date, stock_manager_code, stock_manager_name, reg_date, reg_user)"
	   objBuilder.Append "VALUES('"&emp_no&"','"&stock_level&"','"&emp_name&"','"&emp_company&"',"
	   objBuilder.Append "'"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_in_date&"',"
	   objBuilder.Append "'"&stock_end_date&"','"&emp_no&"','"&emp_name&"',NOW(),'"&user_name&"');"
	Else
		objBuilder.Append "UPDATE met_stock_code SET "
		objBuilder.Append "	stock_name='"&emp_name&"',stock_company='"&emp_company&"',stock_bonbu='"&emp_bonbu&"',stock_saupbu='"&emp_saupbu&"',"
		objBuilder.Append "	stock_team='"&emp_team&"',stock_open_date='"&emp_in_date&"',stock_manager_code='"&emp_no&"',stock_manager_name='"&emp_name&"' "
		objBuilder.Append "WHERE stock_code='"&emp_no&"';"
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rs_stock.Close() : Set rs_stock = Nothing
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 등록되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
'Response.Write "	location.replace('insa_mg.asp');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write"</script>"
Response.End
%>

