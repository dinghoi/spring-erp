<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

emp_yymm=Request.form("emp_yymm1")
view_condi=Request.form("view_condi1")

'response.write(emp_yymm)
'response.write(view_condi)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_bef = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'if view_condi = "전체" then
       Sql = "select * from emp_master where emp_no < '900000' ORDER BY emp_no ASC"
'   else 
'       Sql = "select * from emp_master where (emp_company = '"&view_condi&"') and (emp_no < '900000') ORDER BY emp_no ASC"
'end if   
	   
Rs.Open Sql, Dbconn, 1

Sql = "SELECT * FROM emp_master_month WHERE emp_month = '"&emp_yymm&"'"
Set Rs_bef=Dbconn.Execute(sql)
if Rs_bef.eof then
   do until rs.eof

	emp_no = rs("emp_no")
	emp_name = rs("emp_name")
    emp_ename = rs("emp_ename")
    emp_type = rs("emp_type")
    emp_sex = rs("emp_sex")
    emp_person1 = rs("emp_person1")
    emp_person2 = rs("emp_person2")
    emp_image = rs("emp_image")
	att_file = rs("emp_image")
    emp_first_date = rs("emp_first_date")
    emp_in_date = rs("emp_in_date")
    emp_gunsok_date = rs("emp_gunsok_date")
    emp_yuncha_date = rs("emp_yuncha_date")
    emp_end_gisan = rs("emp_end_gisan")
    emp_end_date = rs("emp_end_date")
	if rs("emp_end_date") = "" or isnull(rs("emp_end_date")) then
           emp_end_date = "1900-01-01"
    end if

    emp_company = rs("emp_company")
    emp_bonbu = rs("emp_bonbu")
    emp_saupbu = rs("emp_saupbu")
    emp_team = rs("emp_team")
    emp_org_code = rs("emp_org_code")
    emp_org_name = rs("emp_org_name")
    emp_org_baldate = rs("emp_org_baldate")
    emp_stay_code = rs("emp_stay_code")
	emp_stay_name = rs("emp_stay_name")
    emp_reside_place = rs("emp_reside_place")
	emp_reside_company = rs("emp_reside_company")
    emp_grade = rs("emp_grade")
    emp_grade_date = rs("emp_grade_date")
    emp_job = rs("emp_job")
    emp_position = rs("emp_position")
    emp_jikgun = rs("emp_jikgun")
    emp_jikmu = rs("emp_jikmu")
    emp_birthday = rs("emp_birthday")
    emp_birthday_id = rs("emp_birthday_id")
    emp_family_zip = rs("emp_family_zip")
    emp_family_sido = rs("emp_family_sido")
    emp_family_gugun = rs("emp_family_gugun")
    emp_family_dong = rs("emp_family_dong")
    emp_family_addr = rs("emp_family_addr")
    emp_zipcode = rs("emp_zipcode")
    emp_sido = rs("emp_sido")
    emp_gugun = rs("emp_gugun")
    emp_dong = rs("emp_dong")
    emp_addr = rs("emp_addr")
    emp_tel_ddd = rs("emp_tel_ddd")
    emp_tel_no1 = rs("emp_tel_no1")
    emp_tel_no2 = rs("emp_tel_no2")
    emp_hp_ddd = rs("emp_hp_ddd")
    emp_hp_no1 = rs("emp_hp_no1")
    emp_hp_no2 = rs("emp_hp_no2")
    emp_email = rs("emp_email")
    emp_military_id = rs("emp_military_id")
    emp_military_date1 = rs("emp_military_date1")
    emp_military_date2 = rs("emp_military_date2")
    emp_military_grade = rs("emp_military_grade")
    emp_military_comm = rs("emp_military_comm")
    emp_hobby = rs("emp_hobby")
    emp_faith = rs("emp_faith")
    emp_last_edu = rs("emp_last_edu")
    emp_marry_date = rs("emp_marry_date")
    emp_disabled = rs("emp_disabled")
    emp_disab_grade = rs("emp_disab_grade")
    emp_sawo_id = rs("emp_sawo_id")
    emp_sawo_date = rs("emp_sawo_date")
    emp_emergency_tel = rs("emp_emergency_tel")
    emp_nation_code = rs("emp_nation_code")
	emp_extension_no = rs("emp_extension_no")
	emp_pay_id = rs("emp_pay_id")
	emp_pay_type = rs("emp_pay_type")
	cost_center = rs("cost_center")
	cost_group = rs("cost_group")
    emp_reg_date = rs("emp_reg_date")
    emp_reg_user = rs("emp_reg_user")
	emp_mod_date = rs("emp_mod_date")
    emp_mod_user = rs("emp_mod_user")
	emp_old_no = rs("emp_old_no")
	if rs("emp_org_baldate") = "" or isnull(rs("emp_org_baldate")) then
           emp_org_baldate = "1900-01-01"
    end if
	if rs("emp_grade_date") = "" or isnull(rs("emp_grade_date")) then
           emp_grade_date = "1900-01-01"
    end if
	if rs("emp_sawo_date") = "" or isnull(rs("emp_sawo_date")) then
           emp_sawo_date = "1900-01-01"
    end if
	if rs("emp_military_date1") = "" or isnull(rs("emp_military_date1")) then
           emp_military_date1 = "1900-01-01"
    end if
	if rs("emp_military_date2") = "" or isnull(rs("emp_military_date2")) then
           emp_military_date2 = "1900-01-01"
    end if
	if rs("emp_marry_date") = "" or isnull(rs("emp_marry_date")) then
           emp_marry_date = "1900-01-01"
    end if
	if rs("emp_birthday") = "" or isnull(rs("emp_birthday")) then
           emp_birthday = "1900-01-01"
    end if
'	if rs("emp_reg_date") = "" or isnull(rs("emp_reg_date")) then
'           emp_reg_date = now()
'    end if
'	if rs("emp_mod_date") = "" or isnull(rs("emp_mod_date")) then
'           emp_mod_date = now()
'    end if
   
	sql = "insert into emp_master_month(emp_month,emp_no,emp_name,emp_ename,emp_type,emp_sex,emp_person1,emp_person2,emp_image,emp_first_date,emp_in_date,emp_gunsok_date,emp_yuncha_date,emp_end_gisan,emp_end_date,emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code,emp_org_name,emp_org_baldate,emp_stay_code,emp_stay_name,emp_reside_place,emp_reside_company,emp_grade,emp_grade_date,emp_job,emp_position,emp_jikgun,emp_jikmu,emp_birthday,emp_birthday_id,emp_family_zip,emp_family_sido,emp_family_gugun,emp_family_dong,emp_family_addr,emp_zipcode,emp_sido,emp_gugun,emp_dong,emp_addr,emp_tel_ddd,emp_tel_no1,emp_tel_no2,emp_hp_ddd,emp_hp_no1,emp_hp_no2,emp_email,emp_military_id,emp_military_date1,emp_military_date2,emp_military_grade,emp_military_comm,emp_hobby,emp_faith,emp_last_edu,emp_marry_date,emp_disabled_yn,emp_disabled,emp_disab_grade,emp_sawo_id,emp_sawo_date,emp_emergency_tel,emp_extension_no,emp_nation_code,emp_pay_id,emp_pay_type,cost_center,cost_group,emp_old_no) values "
	sql = sql +	" ('"&emp_yymm&"','"&emp_no&"','"&emp_name&"','"&emp_ename&"','"&emp_type&"','"&emp_sex&"','"&emp_person1&"','"&emp_person2&"','"&emp_image&"','"&emp_first_date&"','"&emp_in_date&"','"&emp_gunsok_date&"','"&emp_yuncha_date&"','"&emp_end_gisan&"','"&emp_end_date&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_code&"','"&emp_org_name&"','"&emp_org_baldate&"','"&emp_stay_code&"','"&emp_stay_name&"','"&emp_reside_place&"','"&emp_reside_company&"','"&emp_grade&"','"&emp_grade_date&"','"&emp_job&"','"&emp_position&"','"&emp_jikgun&"','"&emp_jikmu&"','"&emp_birthday&"','"&emp_birthday_id&"','"&emp_family_zip&"','"&emp_family_sido&"','"&emp_family_gugun&"','"&emp_family_dong&"','"&emp_family_addr&"','"&emp_zipcode&"','"&emp_sido&"','"&emp_gugun&"','"&emp_dong&"','"&emp_addr&"','"&emp_tel_ddd&"','"&emp_tel_no1&"','"&emp_tel_no2&"','"&emp_hp_ddd&"','"&emp_hp_no1&"','"&emp_hp_no2&"','"&emp_email&"','"&emp_military_id&"','"&emp_military_date1&"','"&emp_military_date2&"','"&emp_military_grade&"','"&emp_military_comm&"','"&emp_hobby&"','"&emp_faith&"','"&emp_last_edu&"','"&emp_marry_date&"','"&emp_disabled_yn&"','"&emp_disabled&"','"&emp_disab_grade&"','"&emp_sawo_id&"','"&emp_sawo_date&"','"&emp_emergency_tel&"','"&emp_extension_no&"','"&emp_nation_code&"','"&emp_pay_id&"','"&emp_pay_type&"','"&cost_center&"','"&cost_group&"','"&emp_old_no&"')"
	
	dbconn.execute(sql)
	   
		Rs.MoveNext()
    loop		
		response.write"<script language=javascript>"
		response.write"alert('인사 마스타 마감처리가 되었습니다...');"		
		response.write"location.replace('insa_master_month_mg.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('이미 마감처리된 내역이 있습니다...');"		
		response.write"location.replace('insa_master_month_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
