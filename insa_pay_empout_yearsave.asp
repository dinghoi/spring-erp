<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

    u_type = request.form("u_type")
    emp_no = request.form("emp_no")
	emp_end_date = request.form("emp_end_date")
	emp_company = request.form("emp_company")
	
	end_year = mid(cstr(emp_end_date),1,4)
	
	sum_give_tot =int(request.form("sum_give_tot"))
	sum_bunus_tot = int(request.form("sum_bunus_tot"))
	sum_tax_no = int(request.form("sum_tax_no"))
	sum_wetax = int(request.form("sum_wetax"))
	sum_epi_amt = int(request.form("sum_epi_amt"))
	sum_longcare_amt = int(request.form("sum_longcare_amt"))
	sum_nhis_amt = int(request.form("sum_nhis_amt"))
	sum_nps_amt = int(request.form("sum_nps_amt"))
	
	total_pay = int(request.form("total_pay"))
	bonin_amt = int(request.form("bonin_amt"))
	total_nhis_amt = int(request.form("total_nhis_amt"))
	
	yaer_income_amt = int(request.form("yaer_income_amt"))
	wife_amt = int(request.form("wife_amt"))

    year_soduk_amt = int(request.form("year_soduk_amt"))
	family_amt = int(request.form("family_amt"))
	sp_incom_amt = int(request.form("sp_incom_amt"))    
	family_age20 = int(request.form("family_age20"))
	family_age60 = int(request.form("family_age60"))    

	year_deduct_hap = int(request.form("year_deduct_hap"))
	year_tax_sp = int(request.form("year_tax_sp"))               

    year_cal_tax = int(request.form("year_cal_tax"))
	just_income_tax = int(request.form("just_income_tax"))
	sum_income_tax = int(request.form("sum_income_tax"))
	add_income_tax = int(request.form("add_income_tax"))            
                                
    year_tax_deduct = int(request.form("year_tax_deduct"))
	just_wetax = int(request.form("just_wetax"))
	add_wetax = int(request.form("add_wetax"))    
                
    re_nhis_month = int(request.form("re_nhis_month"))
	re_nhis_hap = int(request.form("re_nhis_hap"))
	re_longcare_hap = int(request.form("re_longcare_hap"))
	cal_nhis_amt = int(request.form("cal_nhis_amt"))    
    cal_long_amt = int(request.form("cal_long_amt"))   
	
	end_id = "2"
	end_pay_type = "1"          
 
set dbconn = server.CreateObject("adodb.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_ytax = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect
	
Sql = "SELECT * FROM emp_master where emp_no = '"+emp_no+"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_name = rs_emp("emp_name")
		emp_first_date = rs_emp("emp_first_date")
		emp_in_date = rs_emp("emp_in_date")
		emp_end_date = rs_emp("emp_end_date")
		emp_type = rs_emp("emp_type")
		emp_grade = rs_emp("emp_grade")
		emp_position = rs_emp("emp_position")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
		emp_reside_place = rs_emp("emp_reside_place")
		emp_reside_company = rs_emp("emp_reside_company")
		emp_disabled = rs_emp("emp_disabled")
		emp_disab_grade = rs_emp("emp_disab_grade")
   else
		emp_first_date = ""
		emp_in_date = ""
		emp_end_date = ""
		emp_type = ""
		emp_grade = ""
		emp_position = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
		emp_reside_place = ""
		emp_reside_company = ""
		emp_disabled = ""
		emp_disab_grade = ""
end if

dbconn.BeginTrans

sms_msg = emp_no + "-" + emp_name + "- 중도퇴직자 정산이 "

Sql = "SELECT * FROM pay_year_end_tax WHERE end_emp_no = '"+emp_no+"' and end_year = '"+end_year+"' and end_id = '2' and end_pay_type = '1' and end_company = '"+emp_company+"'"
Set Rs_ytax=Dbconn.Execute(sql)
if Rs_ytax.eof then

		sql="insert into pay_year_end_tax (end_year,end_id,end_pay_type,end_emp_no,end_company,end_emp_name,end_bonbu,end_saupbu,end_team,end_org_name,end_org_code,end_total_pay,end_income_deduct,end_income_amt,end_a_self_amt,end_a_wife_amt,end_a_age60_amt,end_a_age20_amt,end_pension_premium,end_c_nhis,end_c_epi,end_c_special,end_deduct_amt,end_tax_base,end_calcu_tax,end_e_taxincome,end_settled_tax,end_settled_wetax,end_add_tax,end_add_wetax,end_payment_tax,end_payment_wetax,end_nhis_month_amt,end_add_nhis,end_add_longcare,end_reg_date,end_reg_user) values ('"&end_year&"','"&end_id&"','"&end_pay_type&"','"&emp_no&"','"&emp_company&"','"&emp_name&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_name&"','"&emp_org_code&"','"&total_pay&"','"&yaer_income_amt&"','"&year_soduk_amt&"','"&bonin_amt&"','"&wife_amt&"','"&family_age60&"','"&family_age20&"','"&sum_nps_amt&"','"&total_nhis_amt&"','"&sum_epi_amt&"','"&sp_incom_amt&"','"&year_deduct_hap&"','"&year_tax_sp&"','"&year_cal_tax&"','"&year_tax_deduct&"','"&just_income_tax&"','"&just_wetax&"','"&add_income_tax&"','"&add_wetax&"','"&sum_income_tax&"','"&sum_wetax&"','"&re_nhis_month&"','"&cal_nhis_amt&"','"&cal_long_amt&"',now(),'"&user_name&"')"
		response.write(sql)
		dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"location.replace('insa_pay_empout_year.asp');"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

  else
	response.write"<script language=javascript>"
	response.write"alert('이미 중도퇴직자 정산처리를 하였습니다..');"		
	'response.write"location.replace('insa_pay_empout_year.asp');"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End
end if	

	dbconn.Close()
	Set dbconn = Nothing

%>
