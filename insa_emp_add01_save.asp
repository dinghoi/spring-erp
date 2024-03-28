<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include file="xmlrpc.asp"-->
<!--#include file="class.EmmaSMS.asp"-->
<%
'	on Error resume next

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

    curr_date = mid(cstr(now()),1,10)

    u_type = abc("u_type")
	
    emp_no = abc("emp_no")
 
    emp_name = abc("emp_name")
    emp_ename = abc("emp_ename")
    emp_type = abc("emp_type")
    emp_sex = abc("emp_sex")
    emp_person1 = abc("emp_person1")
    emp_person2 = abc("emp_person2")
	if emp_person2 <> "" then
	   sex_id = mid(cstr(emp_person2),1,1)
	   if sex_id = "1" then
	         emp_sex = "남"
		  else
		     emp_sex = "여"
	   end if
	end if

    emp_first_date = abc("emp_first_date")
    emp_in_date = abc("emp_in_date")
    emp_gunsok_date = abc("emp_gunsok_date")
    emp_yuncha_date = abc("emp_yuncha_date")
    emp_end_gisan = abc("emp_end_gisan")
    emp_end_date = abc("emp_end_date")
    emp_company = abc("emp_company")
    emp_bonbu = abc("emp_bonbu")
    emp_saupbu = abc("emp_saupbu")
    emp_team = abc("emp_team")
    emp_org_code = abc("emp_org_code")
    emp_org_name = abc("emp_org_name")
	emp_org_baldate = abc("emp_org_baldate")
    emp_stay_code = abc("emp_stay_code")
	emp_stay_name = abc("emp_stay_name")
    emp_reside_place = abc("emp_reside_place")
	emp_reside_company = abc("emp_reside_company")
	
	emp_org_level = abc("emp_org_level")
	if emp_org_level = "상주처" then
	          reside = "1"
	   else 
	          reside = "0"
    end if
    emp_grade = abc("emp_grade")
	emp_grade_date = abc("emp_grade_date")
    emp_job = abc("emp_job")
    emp_position = abc("emp_position")
    emp_jikmu = abc("emp_jikmu")
    emp_birthday = abc("emp_birthday")
    emp_birthday_id = abc("emp_birthday_id")
    emp_family_zip = abc("emp_family_zip")
    emp_family_sido = abc("emp_family_sido")
    emp_family_gugun = abc("emp_family_gugun")
    emp_family_dong = abc("emp_family_dong")
    emp_family_addr = abc("emp_family_addr")
    emp_zipcode = abc("emp_zipcode")
    emp_sido = abc("emp_sido")
    emp_gugun = abc("emp_gugun")
    emp_dong = abc("emp_dong")
    emp_addr = abc("emp_addr")
    emp_tel_ddd = abc("emp_tel_ddd")
    emp_tel_no1 = abc("emp_tel_no1")
    emp_tel_no2 = abc("emp_tel_no2")
    emp_hp_ddd = abc("emp_hp_ddd")
    emp_hp_no1 = abc("emp_hp_no1")
    emp_hp_no2 = abc("emp_hp_no2")
    emp_email = abc("emp_email")
    emp_military_id = abc("emp_military_id")
    emp_military_date1 = abc("emp_military_date1")
    emp_military_date2 = abc("emp_military_date2")
    emp_military_grade = abc("emp_military_grade")
    emp_military_comm = abc("emp_military_comm")
    emp_hobby = abc("emp_hobby")
    emp_faith = abc("emp_faith")
    emp_marry_date = abc("emp_marry_date")
    emp_disabled = abc("emp_disabled")
    emp_disab_grade = abc("emp_disab_grade")
	if emp_disabled = "해당사항없음" or emp_disabled = "" then
	   emp_disabled_yn = "N"
	   emp_disab_grade = ""
	   else 
	   emp_disabled_yn = "Y"
	end if
    emp_sawo_id = abc("emp_sawo_id")
	if emp_sawo_id = "Y" then
	        if u_type = "U" then
	                emp_sawo_date = abc("emp_sawo_date")
			   else
			        emp_sawo_date = abc("emp_in_date")
			end if
	   else
	        emp_sawo_date = "1900-01-01"
	end if		

    emp_emergency_tel = abc("emp_emergency_tel")
	emp_extension_no = abc("emp_extension_no")
	emp_last_edu = abc("emp_last_edu")
	cost_center = abc("cost_center")
	cost_group = abc("cost_group")
	if emp_org_level = "상주처" then
	          cost_center = "상주직접비"
    end if
	if cost_center = "상주직접비" then
	   if isnull(cost_group) or cost_group = "" then
	        cost_group =  emp_reside_company
	   end if
	end if
	mg_group = abc("mg_group")
	emp_pay_id = abc("emp_pay_id")
'	emp_pay_id = "0"
	emp_pay_type = "1"
    emp_nation_code = "001"
	
	kwon_email = emp_email + "@k-won.co.kr"
	emp_hp = emp_hp_ddd + "-" + emp_hp_no1 + "-" + emp_hp_no2
	
	v_att_file= abc("v_att_file")
	
	path = Server.MapPath ("/emp_photo")
	
	Set filenm = abc("att_file")(1)
	filename = filenm
	if filenm <> "" then 
		filename = filenm.safeFileName	
		fileType = mid(filename,inStrRev(filename,".")+1)
		filename = emp_name + "_" + emp_no + "photo." + fileType
		save_path = path & "\" & filename
	end if	
		
	if filenm.length > 1024*1024*8  then 
    	response.write "<script language=javascript>"
      	response.write "alert('파일 용량 2M를 넘으면 안됩니다.');"
		response.write "history.go(-1);"
      	response.write "</script>"
      	response.end
	End If	
	
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
	
	if u_type ="U" then
	   emp_mod_user = request.cookies("nkpmg_user")("coo_user_name")
	   user_id = request.cookies("nkpmg_user")("coo_user_id")
	  else 
       emp_reg_user = request.cookies("nkpmg_user")("coo_user_name")
	   user_id = request.cookies("nkpmg_user")("coo_user_id")
	end if

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_memb = Server.CreateObject("ADODB.Recordset")
	Set rs_sawo = Server.CreateObject("ADODB.Recordset")
	Set rs_stock = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans


if	u_type = "U" then
	
	if filenm <> "" then 
	   filenm.save save_path
    sql = "update emp_master set emp_name ='"+emp_name+"',emp_ename ='"+emp_ename+"',emp_type ='"+emp_type+"',emp_sex ='"+emp_sex
    sql = sql + "',emp_person1 ='"+emp_person1+"',emp_person2 ='"+emp_person2+"',emp_image ='"+filename+"',emp_first_date ='"+emp_first_date+"',emp_in_date ='"+emp_in_date+"',emp_gunsok_date ='"+emp_gunsok_date+"',emp_yuncha_date ='"+emp_yuncha_date
    sql = sql + "',emp_end_gisan ='"+emp_end_gisan+ "',emp_company ='"+emp_company+"',emp_bonbu ='"+emp_bonbu+"',emp_saupbu ='"+emp_saupbu+"',emp_team ='"+emp_team+"',emp_org_code ='"+emp_org_code+"',emp_org_name ='"+emp_org_name+"',emp_grade ='"+emp_grade+"',emp_job ='"+emp_job+"',emp_position ='"+emp_position+"',emp_stay_code ='"+emp_stay_code+"',emp_stay_name ='"+emp_stay_name+"',emp_reside_place ='"+emp_reside_place+"',emp_reside_company ='"+emp_reside_company+"',emp_jikmu ='"+emp_jikmu+"',emp_birthday ='"+emp_birthday+"',emp_birthday_id ='"+emp_birthday_id
    sql = sql + "',emp_family_zip ='"+emp_family_zip+"',emp_family_sido ='"+emp_family_sido+"',emp_family_gugun ='"+emp_family_gugun+"', emp_family_dong ='"+emp_family_dong+"',emp_family_addr ='"+emp_family_addr+"',emp_zipcode ='"+emp_zipcode+"',emp_sido ='"+emp_sido
    sql = sql + "',emp_gugun ='"+emp_gugun+"',emp_dong ='"+emp_dong+"',emp_addr ='"+emp_addr+"',emp_tel_ddd ='"+emp_tel_ddd+"',emp_tel_no1 ='"+emp_tel_no1+"',emp_tel_no2 ='"+emp_tel_no2+"',emp_hp_ddd ='"+emp_hp_ddd+"',emp_hp_no1 ='"+emp_hp_no1+"',emp_hp_no2 ='"+emp_hp_no2
    sql = sql + "',emp_email ='"+emp_email+"',emp_military_id ='"+emp_military_id+"',emp_military_date1 ='"+emp_military_date1+"', emp_military_date2 ='"+emp_military_date2+"',emp_military_grade ='"+emp_military_grade+"',emp_military_comm ='"+emp_military_comm
    sql = sql + "',emp_hobby ='"+emp_hobby+"',emp_faith ='"+emp_faith+"',emp_last_edu ='"+emp_last_edu+"',emp_marry_date ='"+emp_marry_date+"',emp_disabled_yn ='"+emp_disabled_yn+"',emp_disabled ='"+emp_disabled+"',emp_disab_grade ='"+emp_disab_grade+"',emp_sawo_id ='"+emp_sawo_id+"',emp_sawo_date ='"+emp_sawo_date
    sql = sql + "',emp_emergency_tel ='"+emp_emergency_tel+"',emp_extension_no ='"+emp_extension_no+"',cost_center ='"+cost_center+"',cost_group ='"+cost_group+"',emp_mod_user = '"+emp_mod_user+"',emp_mod_date = now() where emp_no ='"+emp_no+"'"
	Else
    sql = "update emp_master set emp_name ='"+emp_name+"',emp_ename ='"+emp_ename+"',emp_type ='"+emp_type+"',emp_sex ='"+emp_sex
    sql = sql + "',emp_person1 ='"+emp_person1+"',emp_person2 ='"+emp_person2+"',emp_first_date ='"+emp_first_date+"',emp_in_date ='"+emp_in_date+"',emp_gunsok_date ='"+emp_gunsok_date+"',emp_yuncha_date ='"+emp_yuncha_date
    sql = sql + "',emp_end_gisan ='"+emp_end_gisan+ "',emp_company ='"+emp_company+"',emp_bonbu ='"+emp_bonbu+"',emp_saupbu ='"+emp_saupbu+"',emp_team ='"+emp_team+"',emp_org_code ='"+emp_org_code+"',emp_org_name ='"+emp_org_name+"',emp_grade ='"+emp_grade+"',emp_job ='"+emp_job+"',emp_position ='"+emp_position+"',emp_stay_code ='"+emp_stay_code+"',emp_stay_name ='"+emp_stay_name+"',emp_reside_place ='"+emp_reside_place+"',emp_reside_company ='"+emp_reside_company+"',emp_jikmu ='"+emp_jikmu+"',emp_birthday ='"+emp_birthday+"',emp_birthday_id ='"+emp_birthday_id
    sql = sql + "',emp_family_zip ='"+emp_family_zip+"',emp_family_sido ='"+emp_family_sido+"',emp_family_gugun ='"+emp_family_gugun+"', emp_family_dong ='"+emp_family_dong+"',emp_family_addr ='"+emp_family_addr+"',emp_zipcode ='"+emp_zipcode+"',emp_sido ='"+emp_sido
    sql = sql + "',emp_gugun ='"+emp_gugun+"',emp_dong ='"+emp_dong+"',emp_addr ='"+emp_addr+"',emp_tel_ddd ='"+emp_tel_ddd+"',emp_tel_no1 ='"+emp_tel_no1+"',emp_tel_no2 ='"+emp_tel_no2+"',emp_hp_ddd ='"+emp_hp_ddd+"',emp_hp_no1 ='"+emp_hp_no1+"',emp_hp_no2 ='"+emp_hp_no2
    sql = sql + "',emp_email ='"+emp_email+"',emp_military_id ='"+emp_military_id+"',emp_military_date1 ='"+emp_military_date1+"', emp_military_date2 ='"+emp_military_date2+"',emp_military_grade ='"+emp_military_grade+"',emp_military_comm ='"+emp_military_comm
    sql = sql + "',emp_hobby ='"+emp_hobby+"',emp_faith ='"+emp_faith+"',emp_last_edu ='"+emp_last_edu+"',emp_marry_date ='"+emp_marry_date+"',emp_disabled_yn ='"+emp_disabled_yn+"',emp_disabled ='"+emp_disabled+"',emp_disab_grade ='"+emp_disab_grade+"',emp_sawo_id ='"+emp_sawo_id+"',emp_sawo_date ='"+emp_sawo_date
    sql = sql + "',emp_emergency_tel ='"+emp_emergency_tel+"',emp_extension_no ='"+emp_extension_no+"',emp_pay_id ='"+emp_pay_id+"',cost_center ='"+cost_center+"',cost_group ='"+cost_group+"',emp_mod_user = '"+emp_mod_user+"',emp_mod_date = now() where emp_no ='"+emp_no+"'"
	end if
	'response.write(sql)  
	dbconn.execute(sql)	
else
	if filenm <> "" then 
	   filenm.save save_path
        sql = "insert into emp_master(emp_no,emp_name,emp_ename,emp_type,emp_sex,emp_person1,emp_person2,emp_image,emp_first_date,emp_in_date,emp_gunsok_date,emp_yuncha_date,emp_end_gisan,emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code,emp_org_name,emp_stay_code,emp_stay_name,emp_reside_place,emp_reside_company,emp_grade,emp_job,emp_position,emp_jikgun,emp_jikmu,emp_birthday,emp_birthday_id,emp_family_zip,emp_family_sido,emp_family_gugun,emp_family_dong,emp_family_addr,emp_zipcode,emp_sido,emp_gugun,emp_dong,emp_addr,emp_tel_ddd,emp_tel_no1,emp_tel_no2,emp_hp_ddd,emp_hp_no1,emp_hp_no2,emp_email,emp_military_id,emp_military_date1,emp_military_date2,emp_military_grade,emp_military_comm,emp_hobby,emp_faith,emp_last_edu,emp_marry_date,emp_disabled_yn,emp_disabled,emp_disab_grade,emp_sawo_id,emp_sawo_date,emp_emergency_tel,emp_extension_no,emp_nation_code,emp_pay_id,emp_pay_type,cost_center,cost_group,emp_reg_date,emp_reg_user) values "
		sql = sql +	" ('"&emp_no&"','"&emp_name&"','"&emp_ename&"','"&emp_type&"','"&emp_sex&"','"&emp_person1&"','"&emp_person2&"','"&filename&"','"&emp_first_date&"','"&emp_in_date&"','"&emp_gunsok_date&"','"&emp_yuncha_date&"','"&emp_end_gisan&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_code&"','"&emp_org_name&"','"&emp_stay_code&"','"&emp_stay_name&"','"&emp_reside_place&"','"&emp_reside_company&"','"&emp_grade&"','"&emp_job&"','"&emp_position&"','"&emp_jikgun&"','"&emp_jikmu&"','"&emp_birthday&"','"&emp_birthday_id&"','"&emp_family_zip&"','"&emp_family_sido&"','"&emp_family_gugun&"','"&emp_family_dong&"','"&emp_family_addr&"','"&emp_zipcode&"','"&emp_sido&"','"&emp_gugun&"','"&emp_dong&"','"&emp_addr&"','"&emp_tel_ddd&"','"&emp_tel_no1&"','"&emp_tel_no2&"','"&emp_hp_ddd&"','"&emp_hp_no1&"','"&emp_hp_no2&"','"&emp_email&"','"&emp_military_id&"','"&emp_military_date1&"','"&emp_military_date2&"','"&emp_military_grade&"','"&emp_military_comm&"','"&emp_hobby&"','"&emp_faith&"','"&emp_last_edu&"','"&emp_marry_date&"','"&emp_disabled_yn&"','"&emp_disabled&"','"&emp_disab_grade&"','"&emp_sawo_id&"','"&emp_sawo_date&"','"&emp_emergency_tel&"','"&emp_extension_no&"','"&emp_nation_code&"','"&emp_pay_id&"','"&emp_pay_type&"','"&cost_center&"','"&cost_group&"',now(),'"&emp_reg_user&"')"
	  Else
        sql = "insert into emp_master(emp_no,emp_name,emp_ename,emp_type,emp_sex,emp_person1,emp_person2,emp_first_date,emp_in_date,emp_gunsok_date,emp_yuncha_date,emp_end_gisan,emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code,emp_org_name,emp_stay_code,emp_stay_name,emp_reside_place,emp_reside_company,emp_grade,emp_job,emp_position,emp_jikgun,emp_jikmu,emp_birthday,emp_birthday_id,emp_family_zip,emp_family_sido,emp_family_gugun,emp_family_dong,emp_family_addr,emp_zipcode,emp_sido,emp_gugun,emp_dong,emp_addr,emp_tel_ddd,emp_tel_no1,emp_tel_no2,emp_hp_ddd,emp_hp_no1,emp_hp_no2,emp_email,emp_military_id,emp_military_date1,emp_military_date2,emp_military_grade,emp_military_comm,emp_hobby,emp_faith,emp_last_edu,emp_marry_date,emp_disabled_yn,emp_disabled,emp_disab_grade,emp_sawo_id,emp_sawo_date,emp_emergency_tel,emp_extension_no,emp_nation_code,emp_pay_id,emp_pay_type,cost_center,cost_group,emp_reg_date,emp_reg_user) values "
		sql = sql +	" ('"&emp_no&"','"&emp_name&"','"&emp_ename&"','"&emp_type&"','"&emp_sex&"','"&emp_person1&"','"&emp_person2&"','"&emp_first_date&"','"&emp_in_date&"','"&emp_gunsok_date&"','"&emp_yuncha_date&"','"&emp_end_gisan&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_code&"','"&emp_org_name&"','"&emp_stay_code&"','"&emp_stay_name&"','"&emp_reside_place&"','"&emp_reside_company&"','"&emp_grade&"','"&emp_job&"','"&emp_position&"','"&emp_jikgun&"','"&emp_jikmu&"','"&emp_birthday&"','"&emp_birthday_id&"','"&emp_family_zip&"','"&emp_family_sido&"','"&emp_family_gugun&"','"&emp_family_dong&"','"&emp_family_addr&"','"&emp_zipcode&"','"&emp_sido&"','"&emp_gugun&"','"&emp_dong&"','"&emp_addr&"','"&emp_tel_ddd&"','"&emp_tel_no1&"','"&emp_tel_no2&"','"&emp_hp_ddd&"','"&emp_hp_no1&"','"&emp_hp_no2&"','"&emp_email&"','"&emp_military_id&"','"&emp_military_date1&"','"&emp_military_date2&"','"&emp_military_grade&"','"&emp_military_comm&"','"&emp_hobby&"','"&emp_faith&"','"&emp_last_edu&"','"&emp_marry_date&"','"&emp_disabled_yn&"','"&emp_disabled&"','"&emp_disab_grade&"','"&emp_sawo_id&"','"&emp_sawo_date&"','"&emp_emergency_tel&"','"&emp_extension_no&"','"&emp_nation_code&"','"&emp_pay_id&"','"&emp_pay_type&"','"&cost_center&"','"&cost_group&"',now(),'"&emp_reg_user&"')"
	 end if	
	 'response.write(sql)
	 dbconn.execute(sql)

end if	 
' 경조회 sawo_mem에 등록	
    if emp_sawo_id = "Y" then
	   sql="select * from emp_sawo_mem where sawo_empno='"&emp_no&"'"
	   set rs_sawo=dbconn.execute(sql)

       if rs_sawo.eof then
  	      sql = "insert into emp_sawo_mem(sawo_empno,sawo_date,sawo_id,sawo_emp_name,sawo_company,sawo_orgcode,sawo_org_name,sawo_target,sawo_in_count,sawo_in_pay,sawo_give_count,sawo_give_pay) values "
		sql = sql +	" ('"&emp_no&"','"&emp_sawo_date&"','입사','"&emp_name&"','"&emp_company&"','"&emp_org_code&"','"&emp_org_name&"','Y',0,0,0,0)"

		  dbconn.execute(sql)	 
	    end if
	end if
	 
' 로긴 memb에 등록	 
if emp_no <> "" or emp_no <> " " then
    sql="select * from memb where user_id='"&emp_no&"'"
	set rs_memb=dbconn.execute(sql)

    if rs_memb.eof then
       sql = "insert into memb(user_id,pass,emp_no,user_name,user_grade,position,emp_company,bonbu,saupbu,team,org_name,hp,email,reside_place,reside_company,reside,mg_group,grade,sms,help_yn,reg_date,reg_id,reg_name) values "
	   sql = sql +	" ('"&emp_no&"','"&emp_person2&"','"&emp_no&"','"&emp_name&"','"&emp_job&"','"&emp_position&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_name&"','"&emp_hp&"','"&kwon_email&"','"&emp_reside_place&"','"&emp_reside_company&"','"&reside&"','"&mg_group&"','4','N','N',now(),'"&user_id&"','"&emp_reg_user&"')"

		'response.write(sql)
		dbconn.execute(sql)	 
	else
	    sql = "update memb set user_name='"&emp_name&"',user_grade='"&emp_job&"',position='"&emp_position&"',emp_company='"&emp_company&"',bonbu='"&emp_bonbu&"',saupbu='"&emp_saupbu&"',team='"&emp_team&"',org_name='"&emp_org_name&"',hp='"&emp_hp&"',email='"&kwon_email&"',reside_place='"&emp_reside_place&"',reside_company='"&emp_reside_company&"',reside='"&reside&"',mg_group='"&mg_group&"',mod_id='"&user_id&"',mod_date=now() where user_id='"&emp_no&"'"

		'response.write sql
		
		dbconn.execute(sql)	  
    end if
end if

' 창고코드 등록	 
if emp_no <> "" or emp_no <> " " then
    sql="select * from met_stock_code where stock_code='"&emp_no&"'"
	set rs_stock=dbconn.execute(sql)

    if rs_stock.eof then
       stock_end_date = "1900-01-01"
	   stock_level = "개인"
	   sql = "insert into met_stock_code (stock_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code,stock_manager_name"
		        sql = sql + ",reg_date,reg_user) values "
		        sql = sql + " ('"&emp_no&"','"&stock_level&"','"&emp_name&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_in_date&"','"&stock_end_date&"','"&emp_no&"','"&emp_name&"',now(),'"&emp_reg_user&"')"        

		'response.write(sql)
		dbconn.execute(sql)	 
	else
	    sql = "update met_stock_code set stock_name='"&emp_name&"',stock_company='"&emp_company&"',stock_bonbu='"&emp_bonbu&"',stock_saupbu='"&emp_saupbu&"',stock_team='"&emp_team&"',stock_open_date='"&emp_in_date&"',stock_manager_code='"&emp_no&"',stock_manager_name='"&emp_name&"' where stock_code='"&emp_no&"'"

		'response.write sql
		
		dbconn.execute(sql)	  
    end if
end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"location.replace('insa_mg.asp');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"			
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>

