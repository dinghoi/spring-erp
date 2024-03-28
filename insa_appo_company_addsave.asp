<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	reg_user = request.cookies("nkpmg_user")("coo_user_name")
	mod_user = request.cookies("nkpmg_user")("coo_user_name")
	user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	curr_date = mid(cstr(now()),1,10)

	app_date = request.form("app_date")
	app_id = request.form("app_id")

	emp_no = request.form("emp_no")
	emp_name = request.form("emp_name")
	
	app_grade = request.form("app_grade")
	app_position = request.form("app_position")
	app_job = request.form("app_job")
	app_to_company = request.form("app_to_company")
	app_to_bonbu = request.form("app_to_bonbu")
	app_to_saupbu = request.form("app_to_saupbu")
	app_to_team = request.form("app_to_team")
	app_org = request.form("app_org")
	app_org_name = request.form("app_org_name")
	
	    sms_msg = emp_no + "-" + emp_name + "- 계열 이동발령"
		new_emp_no = request.form("new_emp_no")
		emp_gunsok_date = request.form("emp_gunsok_date")
        emp_yuncha_date = request.form("emp_yuncha_date")
        emp_end_gisan = request.form("emp_end_gisan")
    	app_be_orgcode = request.form("app_be_orgcode")
	    app_be_org = request.form("app_be_org")
	    app_company = request.form("app_company")
		app_bonbu = request.form("app_bonbu")
		app_saupbu = request.form("app_saupbu")
		app_team = request.form("app_team")
	    app_mv_comment = request.form("app_mv_comment")
		emp_stay_code = request.form("emp_stay_code")
		app_reside_place = request.form("app_reside_place")
		app_reside_company = request.form("app_reside_company")
		stay_name = request.form("stay_name")
		app_jikmu = request.form("emp_jikmu")
        app_org_level = request.form("app_org_level")
	    if app_org_level = "상주처" then
	          reside = "1"
	       else 
	          reside = "0"
        end if
		
		cost_center = request.form("cost_center")
	    cost_group = request.form("app_cost_group")
	    mg_group = request.form("mg_group")
' db update and insert....

	set dbconn = server.CreateObject("adodb.connection")
	
    Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_memb = Server.CreateObject("ADODB.Recordset")
    Set rs_max = Server.CreateObject("ADODB.Recordset")
	Set Rs_sch = Server.CreateObject("ADODB.Recordset")
    Set rs_car = Server.CreateObject("ADODB.Recordset")
    Set rs_qul = Server.CreateObject("ADODB.Recordset")
	Set Rs_fam = Server.CreateObject("ADODB.Recordset")
    Set rs_app = Server.CreateObject("ADODB.Recordset")
    Set rs_edu = Server.CreateObject("ADODB.Recordset")
    Set rs_lan = Server.CreateObject("ADODB.Recordset")
	Set Rs_cmt = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

Sql="select * from emp_master where emp_no = '"&emp_no&"'"
Set rs_emp=DbConn.Execute(Sql)

	emp_name = rs_emp("emp_name")
    emp_ename = rs_emp("emp_ename")
    emp_type = rs_emp("emp_type")
    emp_sex = rs_emp("emp_sex")
    emp_person1 = rs_emp("emp_person1")
    emp_person2 = rs_emp("emp_person2")
    emp_image = rs_emp("emp_image")
    emp_first_date = rs_emp("emp_first_date")
    emp_in_date = app_date
    
    emp_end_date = "1900-01-01"
    emp_org_baldate = app_date

    emp_grade = rs_emp("emp_grade")
    emp_grade_date = rs_emp("emp_grade_date")
    emp_job = rs_emp("emp_job")
    emp_position = rs_emp("emp_position")
    emp_jikgun = rs_emp("emp_jikgun")
    emp_birthday = rs_emp("emp_birthday")
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
    emp_military_date1 = rs_emp("emp_military_date1")
    emp_military_date2 = rs_emp("emp_military_date2")
    emp_military_grade = rs_emp("emp_military_grade")
    emp_military_comm = rs_emp("emp_military_comm")
    emp_hobby = rs_emp("emp_hobby")
    emp_faith = rs_emp("emp_faith")
    emp_last_edu = rs_emp("emp_last_edu")
    emp_marry_date = rs_emp("emp_marry_date")
	emp_disabled_yn = rs_emp("emp_disabled_yn")
    emp_disabled = rs_emp("emp_disabled")
    emp_disab_grade = rs_emp("emp_disab_grade")
    emp_sawo_id = rs_emp("emp_sawo_id")
    emp_sawo_date = rs_emp("emp_sawo_date")
    emp_emergency_tel = rs_emp("emp_emergency_tel")
    emp_pay_id = rs_emp("emp_pay_id")
	emp_extension_no = rs_emp("emp_extension_no")
	emp_old_no = rs_emp("emp_old_no")
'	cost_center = rs("cost_center")
'	cost_group = rs("cost_group")
	
    kwon_email = emp_email + "@k-won.co.kr"

    if emp_birthday = "" or isnull(emp_birthday) then
	   emp_birthday = "1900-01-01"
	end if
	if emp_end_date = "" or isnull(emp_end_date) then
	   emp_end_date = "1900-01-01"
	end if
	if emp_org_baldate = "" or isnull(emp_org_baldate) then
	   emp_org_baldate = "1900-01-01"
	end if
	if emp_grade_date = "" or isnull(emp_grade_date) then
	   emp_grade_date = "1900-01-01"
	end if
	if emp_military_date1 = "" or isnull(emp_military_date1) then
	   emp_military_date1 = "1900-01-01"
	end if
	if emp_military_date2 = "" or isnull(emp_military_date2) then
	   emp_military_date2 = "1900-01-01"
	end if
	if emp_marry_date = "" or isnull(emp_marry_date) then
	   emp_marry_date = "1900-01-01"
	end if
	if emp_sawo_date = "" or isnull(emp_sawo_date) then
	   emp_sawo_date = "1900-01-01"
	end if

	dbconn.BeginTrans


sql = "insert into emp_master(emp_no,emp_name,emp_ename,emp_type,emp_sex,emp_person1,emp_person2,emp_image,emp_first_date,emp_in_date,emp_gunsok_date,emp_yuncha_date,emp_end_gisan,emp_end_date,emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code,emp_org_name,emp_org_baldate,emp_stay_code,emp_stay_name,emp_reside_place,emp_reside_company,emp_grade,emp_grade_date,emp_job,emp_position,emp_jikgun,emp_jikmu,emp_birthday,emp_birthday_id,emp_family_zip,emp_family_sido,emp_family_gugun,emp_family_dong,emp_family_addr,emp_zipcode,emp_sido,emp_gugun,emp_dong,emp_addr,emp_tel_ddd,emp_tel_no1,emp_tel_no2,emp_hp_ddd,emp_hp_no1,emp_hp_no2,emp_email,emp_military_id,emp_military_date1,emp_military_date2,emp_military_grade,emp_military_comm,emp_hobby,emp_faith,emp_last_edu,emp_marry_date,emp_disabled_yn,emp_disabled,emp_disab_grade,emp_sawo_id,emp_sawo_date,emp_emergency_tel,emp_extension_no,emp_nation_code,emp_pay_id,emp_reg_date,emp_reg_user,emp_old_no,cost_group,cost_center) values "
		sql = sql +	" ('"&new_emp_no&"','"&emp_name&"','"&emp_ename&"','"&emp_type&"','"&emp_sex&"','"&emp_person1&"','"&emp_person2&"','"&emp_image&"','"&emp_first_date&"','"&app_date&"','"&emp_gunsok_date&"','"&emp_yuncha_date&"','"&emp_end_gisan&"','"&emp_end_date&"','"&app_company&"','"&app_bonbu&"','"&app_saupbu&"','"&app_team&"','"&app_be_orgcode&"','"&app_be_org&"','"&emp_org_baldate&"','"&emp_stay_code&"','"&stay_name&"','"&app_reside_place&"','"&app_reside_company&"','"&emp_grade&"','"&emp_grade_date&"','"&emp_job&"','"&emp_position&"','"&emp_jikgun&"','"&app_jikmu&"','"&emp_birthday&"','"&emp_birthday_id&"','"&emp_family_zip&"','"&emp_family_sido&"','"&emp_family_gugun&"','"&emp_family_dong&"','"&emp_family_addr&"','"&emp_zipcode&"','"&emp_sido&"','"&emp_gugun&"','"&emp_dong&"','"&emp_addr&"','"&emp_tel_ddd&"','"&emp_tel_no1&"','"&emp_tel_no2&"','"&emp_hp_ddd&"','"&emp_hp_no1&"','"&emp_hp_no2&"','"&emp_email&"','"&emp_military_id&"','"&emp_military_date1&"','"&emp_military_date2&"','"&emp_military_grade&"','"&emp_military_comm&"','"&emp_hobby&"','"&emp_faith&"','"&emp_last_edu&"','"&emp_marry_date&"','"&emp_disabled_yn&"','"&emp_disabled&"','"&emp_disab_grade&"','"&emp_sawo_id&"','"&emp_sawo_date&"','"&emp_emergency_tel&"','"&emp_extension_no&"','"&emp_nation_code&"','0',now(),'"&reg_user&"','"&emp_no&"','"&cost_group&"','"&cost_center&"')"

	 'response.write(sql)
	 dbconn.execute(sql)
 
' 로긴 memb에 등록	 
    sql="select * from memb where user_id='"&new_emp_no&"'"
	set rs_memb=dbconn.execute(sql)

    if rs_memb.eof then
       sql = "insert into memb(user_id,pass,emp_no,user_name,user_grade,position,emp_company,bonbu,saupbu,team,org_name,hp,email,reside_place,reside_company,reside,mg_group,grade,sms,help_yn,reg_date,reg_id,reg_name) values "
	   sql = sql +	" ('"&new_emp_no&"','"&emp_person2&"','"&new_emp_no&"','"&emp_name&"','"&emp_job&"','"&emp_position&"','"&app_company&"','"&app_bonbu&"','"&app_saupbu&"','"&app_team&"','"&app_be_org&"','"&emp_hp&"','"&kwon_email&"','"&app_reside_place&"','"&app_reside_company&"','"&reside&"','"&mg_group&"','4','N','N',now(),'"&user_id&"','"&reg_user&"')"

		dbconn.execute(sql)	 
     end if

    sql="select * from memb where user_id='"&emp_no&"'"
	set rs_memb=dbconn.execute(sql)

    if not rs_memb.eof then
	    sql = "update memb set grade='6' where user_id = '"&emp_no&"'"
				
		dbconn.execute(sql)	  
    end if
	
'가족자료 변경
sql = "select * from emp_family where family_empno = '"&emp_no&"'"
set Rs_fam=dbconn.execute(sql)
if not Rs_fam.eof then
   do until Rs_fam.eof
      family_empno = Rs_fam("family_empno")
	  family_seq = Rs_fam("family_seq")
	  family_rel = Rs_fam("family_rel")
      family_name = Rs_fam("family_name")
      family_birthday = Rs_fam("family_birthday")
      family_birthday_id = Rs_fam("family_birthday_id")
      family_job = Rs_fam("family_job")
      family_live = Rs_fam("family_live")
      family_person1 = Rs_fam("family_person1")
      family_person2 = Rs_fam("family_person2")
	  family_tel_ddd = Rs_fam("family_tel_ddd")
      family_tel_no1 = Rs_fam("family_tel_no1")
      family_tel_no2 = Rs_fam("family_tel_no2")
	  family_support_yn = Rs_fam("family_support_yn")
	  family_reg_date = Rs_fam("family_reg_date")
	  family_reg_user = Rs_fam("family_reg_user")
	  if family_reg_date = "" or isnull(family_reg_date) then
	     family_reg_date = "1900-01-01"
	  end if
	  
      sql = "insert into emp_family (family_empno,family_seq,family_rel,family_name,family_birthday,family_birthday_id,family_job,family_live,family_support_yn,family_person1,family_person2,family_tel_ddd,family_tel_no1,family_tel_no2,family_reg_date,family_reg_user,family_mod_date,family_mod_user) values "
	  sql = sql +	" ('"&new_emp_no&"','"&family_seq&"','"&family_rel&"','"&family_name&"','"&family_birthday&"','"&family_birthday_id&"','"&family_job&"','"&family_live&"','"&family_support_yn&"','"&family_person1&"','"&family_person2&"','"&family_tel_ddd&"','"&family_tel_no1&"','"&family_tel_no2&"',now(),'"&reg_user&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    Rs_fam.MoveNext()
   loop		
end if

'학력사항
Sql="select * from emp_school where sch_empno = '"&emp_no&"'"
Set Rs_sch=DbConn.Execute(Sql)
if not Rs_sch.eof then
   do until Rs_sch.eof
    sch_empno = Rs_sch("sch_empno")
    sch_seq = Rs_sch("sch_seq")
	sch_start_date = Rs_sch("sch_start_date")
    sch_end_date = Rs_sch("sch_end_date")
    sch_dept = Rs_sch("sch_dept")
    sch_major = Rs_sch("sch_major")
    sch_sub_major = Rs_sch("sch_sub_major")
    sch_degree = Rs_sch("sch_degree")
	sch_finish = Rs_sch("sch_finish")
	sch_comment = Rs_sch("sch_comment")
    view_condi = Rs_sch("sch_comment")
	if view_condi = "1" then 
	        sch_school_name = Rs_sch("sch_school_name")
	   else
	        sch_school_name = Rs_sch("sch_school_name")
	end if
	sch_reg_date = Rs_sch("sch_reg_date")
	sch_reg_user = Rs_sch("sch_reg_user")
	if sch_reg_date = "" or isnull(sch_reg_date) then
	     sch_reg_date = "1900-01-01"
	  end if

    sql = "insert into emp_school (sch_empno,sch_seq,sch_start_date,sch_end_date,sch_school_name,sch_dept,sch_major,sch_sub_major,sch_degree,sch_finish,sch_comment,sch_reg_date,sch_reg_user,sch_mod_date,sch_mod_user) values "
	sql = sql +	" ('"&new_emp_no&"','"&sch_seq&"','"&sch_start_date&"','"&sch_end_date&"','"&sch_school_name&"','"&sch_dept&"','"&sch_major&"','"&sch_sub_major&"','"&sch_degree&"','"&sch_finish&"','"&sch_comment&"',now(),'"&reg_user&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    Rs_sch.MoveNext()
   loop		
end if

'경력사항
Sql="select * from emp_career where career_empno = '"&emp_no&"'"
Set rs_car=DbConn.Execute(Sql)
if not rs_car.eof then
   do until rs_car.eof
    career_empno = rs_car("career_empno")
    career_seq = rs_car("career_seq")
	career_join_date = rs_car("career_join_date")
    career_end_date = rs_car("career_end_date")
    career_office = rs_car("career_office")
    career_dept = rs_car("career_dept")
    career_position = rs_car("career_position")
    career_task = rs_car("career_task")
	career_reg_date = rs_car("career_reg_date")
	career_reg_user = rs_car("career_reg_user")
	if career_reg_date = "" or isnull(career_reg_date) then
	     career_reg_date = "1900-01-01"
	end if

    sql = "insert into emp_career(career_empno,career_seq,career_join_date,career_end_date,career_office,career_dept,career_position,career_task,career_reg_date,career_reg_user,career_mod_date,career_mod_user) values "
	sql = sql +	" ('"&new_emp_no&"','"&career_seq&"','"&career_join_date&"','"&career_end_date&"','"&career_office&"','"&career_dept&"','"&career_position&"','"&career_task&"',now(),'"&reg_user&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    rs_car.MoveNext()
   loop		
end if

'자격사항
Sql="select * from emp_qual where qual_empno = '"&emp_no&"'"
Set rs_qul=DbConn.Execute(Sql)
if not rs_qul.eof then
   do until rs_qul.eof
    qual_empno = rs_qul("qual_empno")
    qual_seq = rs_qul("qual_seq")
	qual_type = rs_qul("qual_type")
    qual_grade = rs_qul("qual_grade")
    qual_pass_date = rs_qul("qual_pass_date")
    qual_org = rs_qul("qual_org")
    qual_no = rs_qul("qual_no")
	qual_passport = rs_qul("qual_passport")
	qual_pay_id = rs_qul("qual_pay_id")
	qual_reg_date = rs_qul("qual_reg_date")
	qual_reg_user = rs_qul("qual_reg_user")
	if qual_reg_date = "" or isnull(qual_reg_date) then
	     qual_reg_date = "1900-01-01"
	end if

    sql = "insert into emp_qual(qual_empno,qual_seq,qual_type,qual_grade,qual_pass_date,qual_org,qual_no,qual_passport,qual_pay_id,qual_reg_date,qual_reg_user,qual_mod_date,qual_mod_user) values "
	sql = sql +	" ('"&new_emp_no&"','"&qual_seq&"','"&qual_type&"','"&qual_grade&"','"&qual_pass_date&"','"&qual_org&"','"&qual_no&"','"&qual_passport&"','"&qual_pay_id&"',now(),'"&reg_user&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    rs_qul.MoveNext()
   loop		
end if

'교육사항
Sql="select * from emp_edu where edu_empno = '"&emp_no&"'"
Set rs_edu=DbConn.Execute(Sql)
if not rs_edu.eof then
   do until rs_edu.eof
    edu_empno = rs_edu("edu_empno")
    edu_seq = rs_edu("edu_seq")
	edu_name = rs_edu("edu_name")
    edu_office = rs_edu("edu_office")
    edu_finish_no = rs_edu("edu_finish_no")
    edu_start_date = rs_edu("edu_start_date")
    edu_end_date = rs_edu("edu_end_date")
    edu_pay = rs_edu("edu_pay")
    edu_comment = rs_edu("edu_comment")
    edu_reg_date = rs_edu("edu_reg_date")
	edu_reg_user = rs_edu("edu_reg_user")
	if edu_reg_date = "" or isnull(edu_reg_date) then
	     edu_reg_date = "1900-01-01"
	end if

    sql = "insert into emp_edu (edu_empno,edu_seq,edu_name,edu_office,edu_finish_no,edu_start_date,edu_end_date,edu_pay,edu_comment,edu_reg_date,edu_reg_user,edu_mod_date,edu_mod_user) values "
	sql = sql +	" ('"&new_emp_no&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"','"&edu_start_date&"','"&edu_end_date&"','"&edu_pay&"','"&edu_comment&"',now(),'"&reg_user&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    rs_edu.MoveNext()
   loop		
end if

'어학사항
Sql="select * from emp_language where lang_empno = '"&emp_no&"'"
Set rs_lan=DbConn.Execute(Sql)
if not rs_lan.eof then
   do until rs_lan.eof
    lang_empno = rs_lan("lang_empno")
    lang_seq = rs_lan("lang_seq")
	lang_id = rs_lan("lang_id")
    lang_id_type = rs_lan("lang_id_type")
    lang_point = rs_lan("lang_point")
    lang_grade = rs_lan("lang_grade")
    lang_get_date = rs_lan("lang_get_date")
	lang_reg_date = rs_lan("lang_reg_date")
	lang_reg_user = rs_lan("lang_reg_user")
	if lang_reg_date = "" or isnull(lang_reg_date) then
	     lang_reg_date = "1900-01-01"
	end if

    sql = "insert into emp_language (lang_empno,lang_seq,lang_id,lang_id_type,lang_point,lang_grade,lang_get_date,lang_reg_date,lang_reg_user,lang_mod_date,lang_mod_user) values "
	sql = sql +	" ('"&new_emp_no&"','"&lang_seq&"','"&lang_id&"','"&lang_id_type&"','"&lang_point&"','"&lang_grade&"','"&lang_get_date&"',now(),'"&reg_user&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    rs_lan.MoveNext()
   loop		
end if

'인사특이사항사항
Sql="select * from emp_comment where cmt_empno = '"&emp_no&"'"
Set Rs_cmt=DbConn.Execute(Sql)
if not Rs_cmt.eof then
   do until Rs_cmt.eof
    cmt_empno = Rs_cmt("cmt_empno")
    cmt_date = Rs_cmt("cmt_date")
	cmt_emp_name = Rs_cmt("cmt_emp_name")
    cmt_company = Rs_cmt("cmt_company")
	cmt_bonbu = Rs_cmt("cmt_bonbu")
    cmt_saupbu = Rs_cmt("cmt_saupbu")
	cmt_team = Rs_cmt("cmt_team")
    cmt_org_name = Rs_cmt("cmt_org_name")
	cmt_org_code = Rs_cmt("cmt_org_code")
	cmt_comment = Rs_cmt("cmt_comment")
	cmt_reg_date = Rs_cmt("cmt_reg_date")
	cmt_reg_user = Rs_cmt("cmt_reg_user")
	if cmt_reg_date = "" or isnull(cmt_reg_date) then
	     cmt_reg_date = "1900-01-01"
	end if

    sql = "insert into emp_comment (cmt_empno,cmt_date,cmt_emp_name,cmt_company,cmt_bonbu,cmt_saupbu,cmt_team,cmt_org_name,cmt_org_code,cmt_comment,cmt_reg_date,cmt_reg_user) values "
	sql = sql +	" ('"&new_emp_no&"','"&cmt_date&"','"&cmt_emp_name&"','"&cmt_company&"','"&cmt_bonbu&"','"&cmt_saupbu&"','"&cmt_team&"','"&cmt_org_name&"','"&cmt_org_code&"','"&cmt_comment&"',now(),'"&reg_user&"')"
		dbconn.execute(sql)

	    Rs_cmt.MoveNext()
   loop		
end if


' url = "as_list_ce.asp?page="+page+"&view_sort="+view_sort
  url = "insa_appoint_company.asp"
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "등록되었습니다...."
	end if
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"alert('등록 완료 되었습니다....');"		
	response.write"location.replace('"&url&"');"
'	response.write"history.go(-2);"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
