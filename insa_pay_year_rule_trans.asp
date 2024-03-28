<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

'년도가 바뀌면 급여처리를 위한 당연도 4대보험요율/근로소득세율/개인 연봉파일에 대해 당연도 데이타를 생성하고 1월 급여처리를 해야함

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = now()
be_date = dateadd("yyyy",-1,curr_date)
be_year = cstr(mid(be_date,1,4))

af_year = cstr(mid(curr_date,1,4))

be_year = "2015"
af_year = "2016"
'response.write(be_year)
'response.write("/")
'response.write(af_year)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_amt = Server.CreateObject("ADODB.Recordset")
Set Rs_rule = Server.CreateObject("ADODB.Recordset")
Set Rs_insu = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Dbconn.BeginTrans 

'sql = "delete from pay_income_rule where inc_yyyy ='"&af_year&"'"       '근로소득세율
'    dbconn.execute(sql)
	
'sql = "delete from pay_insurance where insu_yyyy ='"&af_year&"'"        '4대보험 기준등급
'    dbconn.execute(sql)

'sql = "delete from pay_income_amount where inc_yyyy ='"&af_year&"'"     '근로소득간이세액
'    dbconn.execute(sql)
	
'sql = "delete from pay_year_income where incom_year ='"&af_year&"'"     '직원연봉
'    dbconn.execute(sql)	

' 근로소득세율설정
Sql = "SELECT * FROM pay_income_rule where rule_yyyy = '"+be_year+"'"
Rs_rule.Open Sql, Dbconn, 1
while not Rs_rule.eof
       
       rule_id = Rs_rule("rule_id")
	   rule_cl = Rs_rule("rule_cl")
       rule_id_name = Rs_rule("rule_id_name")
	   rule_year_pay =Rs_rule("rule_year_pay")
       rule_st_deduct = Rs_rule("rule_st_deduct")
       rule_exceed = Rs_rule("rule_exceed")
       rule_exceed_rate = Rs_rule("rule_exceed_rate")
       rule_add = Rs_rule("rule_add")
       rule_add_rate = Rs_rule("rule_add_rate")
       rule_comment = Rs_rule("rule_comment")
	   
	   sql="insert into pay_income_rule (rule_yyyy,rule_id,rule_cl,rule_id_name,rule_year_pay,rule_st_deduct,rule_exceed,rule_add,rule_exceed_rate,rule_add_rate,rule_comment,rule_reg_user,rule_reg_date) values ('"&af_year&"','"&rule_id&"','"&rule_cl&"','"&rule_id_name&"','"&rule_year_pay&"','"&rule_st_deduct&"','"&rule_exceed&"','"&rule_add&"','"&rule_exceed_rate&"','"&rule_add_rate&"','"&rule_comment&"','"&emp_user&"',now())"
	   
	   dbconn.execute(sql)
	   
	   Rs_rule.movenext()
Wend
Rs_rule.close()	

' 4대보험 기준등급
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"+be_year+"'"
Rs_insu.Open Sql, Dbconn, 1
while not Rs_insu.eof
       
       insu_id = Rs_insu("insu_id")
	   insu_class = Rs_insu("insu_class")
       insu_id_name = Rs_insu("insu_id_name")
	   from_amt =Rs_insu("from_amt")
       to_amt = Rs_insu("to_amt")
       st_amt = Rs_insu("st_amt")
       tot_rate = Rs_insu("hap_rate")
       emp_rate = Rs_insu("emp_rate")
       com_rate = Rs_insu("com_rate")
       insu_comment = Rs_insu("insu_comment")
	   
	   sql="insert into pay_insurance (insu_yyyy,insu_id,insu_class,insu_id_name,from_amt,to_amt,st_amt,hap_rate,emp_rate,com_rate,insu_comment,reg_user,reg_date) values ('"&af_year&"','"&insu_id&"','"&insu_class&"','"&insu_id_name&"','"&from_amt&"','"&to_amt&"','"&st_amt&"','"&tot_rate&"','"&emp_rate&"','"&com_rate&"','"&insu_comment&"','"&emp_user&"',now())"
	   
	   dbconn.execute(sql)
	   
	   Rs_insu.movenext()
Wend
Rs_insu.close()	

' 근로소득간이세액
Sql = "SELECT * FROM pay_income_amount where inc_yyyy = '"+be_year+"'"
Rs_amt.Open Sql, Dbconn, 1
while not Rs_amt.eof
       
       inc_seq = Rs_amt("inc_seq")
	   inc_from_amt = Rs_amt("inc_from_amt")
       inc_to_amt = Rs_amt("inc_to_amt")
	   inc_st_amt =Rs_amt("inc_st_amt")
       inc_incom1 = Rs_amt("inc_incom1")
       inc_incom2 = Rs_amt("inc_incom2")
       inc_incom3 = Rs_amt("inc_incom3")
       inc_incom4 = Rs_amt("inc_incom4")
       inc_incom5 = Rs_amt("inc_incom5")
	   inc_incom6 = Rs_amt("inc_incom6")
	   inc_incom7 = Rs_amt("inc_incom7")
	   inc_incom8 = Rs_amt("inc_incom8")
	   inc_incom9 = Rs_amt("inc_incom9")
	   inc_incom10 = Rs_amt("inc_incom10")
	   inc_incom11 = Rs_amt("inc_incom11")
	   inc_incom12 = Rs_amt("inc_incom12")
	   if isnull(inc_incom12) or inc_incom12 = "" then
	      inc_incom12 = 0
	   end if
       inc_comment = Rs_amt("inc_comment")
	   
	   sql="insert into pay_income_amount (inc_yyyy,inc_seq,inc_from_amt,inc_to_amt,inc_st_amt,inc_incom1,inc_incom2,inc_incom3,inc_incom4,inc_incom5,inc_incom6,inc_incom7,inc_incom8,inc_incom9,inc_incom10,inc_incom11,inc_incom12,inc_comment,inc_reg_user,inc_reg_date) values ('"&af_year&"','"&inc_seq&"','"&inc_from_amt&"','"&inc_to_amt&"','"&inc_st_amt&"','"&inc_incom1&"','"&inc_incom2&"','"&inc_incom3&"','"&inc_incom4&"','"&inc_incom5&"','"&inc_incom6&"','"&inc_incom7&"','"&inc_incom8&"','"&inc_incom9&"','"&inc_incom10&"','"&inc_incom11&"','"&inc_incom12&"','"&inc_comment&"','"&emp_user&"',now())"
	   
	   dbconn.execute(sql)
	   
	   Rs_amt.movenext()
Wend
Rs_amt.close()	

' 직원연봉
Sql = "SELECT * FROM pay_year_income where incom_year = '"+be_year+"'"
Rs_year.Open Sql, Dbconn, 1
while not Rs_year.eof
       
	   incom_emp_no = Rs_year("incom_emp_no")
	   incom_emp_name = Rs_year("incom_emp_name")
	   incom_in_date = Rs_year("incom_in_date")
	   incom_grade = Rs_year("incom_grade")
	   incom_emp_type = Rs_year("incom_emp_type")
	   if Rs_year("incom_pay_type") = "1" then 
	         incom_pay_type = "근로소득"
	      else
	         incom_pay_type = "사업소득"	
       end if  
	   incom_company = Rs_year("incom_company")
	   incom_org_code = Rs_year("incom_org_code")
	   incom_org_name = Rs_year("incom_org_name")
	
	   incom_base_pay = Rs_year("incom_base_pay")
       incom_overtime_pay = Rs_year("incom_overtime_pay")
       incom_meals_pay = Rs_year("incom_meals_pay")
       incom_severance_pay = Rs_year("incom_severance_pay")
	   incom_total_pay = Rs_year("incom_total_pay")
	   incom_first3_percent = Rs_year("incom_first3_percent")
	   incom_month_amount = Rs_year("incom_month_amount")
	   incom_nps_amount = Rs_year("incom_nps_amount")
	   incom_nhis_amount = Rs_year("incom_nhis_amount")
	   incom_family_cnt = Rs_year("incom_family_cnt")
	   incom_nps = Rs_year("incom_nps")
       incom_nhis = Rs_year("incom_nhis")
       incom_go_yn = Rs_year("incom_go_yn")
       incom_san_yn = Rs_year("incom_san_yn")
       incom_long_yn = Rs_year("incom_long_yn")
       incom_incom_yn = Rs_year("incom_incom_yn")
       incom_wife_yn = Rs_year("incom_wife_yn")
       incom_age20 = Rs_year("incom_age20")
       incom_age60 = Rs_year("incom_age60")
       incom_old = Rs_year("incom_old")
       incom_disab = Rs_year("incom_disab")
       incom_woman = Rs_year("incom_woman")
	   incom_retirement_bank = Rs_year("incom_retirement_bank") 
	if isnull(incom_go_yn) or incom_go_yn = "" then
	      incom_go_yn = "여"
	end if
	if isnull(incom_san_yn) or incom_san_yn = "" then
	      incom_san_yn = "여"
	end if
	if isnull(incom_long_yn) or incom_long_yn = "" then
	      incom_long_yn = "여"
	end if
	if isnull(incom_incom_yn) or incom_incom_yn = "" then
	      incom_incom_yn = "부"
	end if
	if isnull(incom_wife_yn) or incom_wife_yn = "" then
	      incom_wife_yn = "0"
	end if
	if isnull(incom_woman) or incom_woman = "" then
	      incom_woman = "0"
	end if
	   
	   sql = "insert into pay_year_income (incom_emp_no,incom_year,incom_emp_name,incom_in_date,incom_grade,incom_emp_type,incom_pay_type,incom_company,incom_org_code,incom_org_name,incom_base_pay,incom_overtime_pay,incom_meals_pay,incom_severance_pay,incom_total_pay,incom_first3_percent,incom_month_amount,incom_nps_amount,incom_nhis_amount,incom_family_cnt,incom_nps,incom_nhis,incom_go_yn,incom_san_yn,incom_long_yn,incom_incom_yn,incom_wife_yn,incom_age20,incom_age60,incom_old,incom_disab,incom_woman,incom_retirement_bank,incom_reg_date,incom_reg_user) values "
		sql = sql +	" ('"&incom_emp_no&"','"&af_year&"','"&incom_emp_name&"','"&incom_in_date&"','"&incom_grade&"','"&incom_emp_type&"','"&incom_pay_type&"','"&incom_company&"','"&incom_org_code&"','"&incom_org_name&"','"&incom_base_pay&"','"&incom_overtime_pay&"','"&incom_meals_pay&"','"&incom_severance_pay&"','"&incom_total_pay&"','"&incom_first3_percent&"','"&incom_month_amount&"','"&incom_nps_amount&"','"&incom_nhis_amount&"','"&incom_family_cnt&"','"&incom_nps&"','"&incom_nhis&"','"&incom_go_yn&"','"&incom_san_yn&"','"&incom_long_yn&"','"&incom_incom_yn&"','"&incom_wife_yn&"','"&incom_age20&"','"&incom_age60&"','"&incom_old&"','"&incom_disab&"','"&incom_woman&"','"&incom_retirement_bank&"',now(),'"&emp_user&"')"
	   
	   dbconn.execute(sql)
	   
	   Rs_year.movenext()
Wend
Rs_year.close()	
	
if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	response.write"<script language=javascript>"
	response.write"alert('"&af_year&"...급여 기초자료 이월처리가 되었습니다...');"		
	'response.write"location.replace('insa_master_month_mg.asp');"
	response.write"location.replace('insa_pay_rule_mg.asp');"
	response.write"</script>"
	Response.End
end if

dbconn.Close()
Set dbconn = Nothing
	
%>
