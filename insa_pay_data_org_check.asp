<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

'SELECT [2015도로명주소_all].building
'FROM 2015도로명주소_all
'WHERE isnull([2015도로명주소_all].building);

'UPDATE 2015도로명주소_all SET [2015도로명주소_all].dong = "" where isnull([2015도로명주소_all].dong);


'UPDATE emp_master_month INNER JOIN o_emp03 ON (emp_master_month.emp_month = [o_emp03].emp_month) and (emp_master_month.emp_no = [o_emp03].emp_no) SET emp_master_month.emp_org_code = [o_emp03].emp_org_code, emp_master_month.emp_stay_name = [o_emp03].emp_stay_name;

'UPDATE emp_master INNER JOIN o_emp05 ON (emp_master.emp_no = [o_emp05].emp_no) SET emp_master.emp_stay_name = [o_emp05].emp_stay_name;

'UPDATE emp_org_mst SET org_date = '2015-04-01' where org_company = '코리아디엔씨';

'UPDATE emp_sawo_mem INNER JOIN emp_master ON (emp_sawo_mem.sawo_empno = [emp_master].emp_no) SET emp_sawo_mem.sawo_orgcode = [emp_master].emp_org_code, emp_sawo_mem.sawo_company = [emp_master].emp_company, emp_sawo_mem.sawo_org_name = [emp_master].emp_org_name;

'UPDATE car_info INNER JOIN emp_master ON (car_info.owner_emp_no = [emp_master].emp_no) SET car_info.car_use_dept = [emp_master].emp_org_name;


emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

pmg_yymm="201507"
'view_condi = "케이원정보통신"
'view_condi = "에스유에이치"
'view_condi = "케이네트웍스"
'view_condi = "휴디스"
'view_condi = "코리아디엔씨"
pmg_id = "4" '1 2 4
pmg_date = "2015-07-31"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

i = 0
j = 0

'Sql = "select * from pay_month_give where (pmg_yymm > '201301' )  ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    j = j + 1
	emp_no = rs("pmg_emp_no")
	
	'pmg_yymm = rs("pmg_yymm")
	'pmg_id = rs("pmg_id")
	view_condi = rs("pmg_company")
	
	pmg_base_pay = rs("pmg_base_pay")
	pmg_meals_pay = rs("pmg_meals_pay")
	pmg_postage_pay = rs("pmg_postage_pay")
	pmg_re_pay = rs("pmg_re_pay")
	pmg_overtime_pay = rs("pmg_overtime_pay")
	pmg_car_pay = rs("pmg_car_pay")
	pmg_position_pay = rs("pmg_position_pay")
	pmg_custom_pay = rs("pmg_custom_pay")
	pmg_job_pay = rs("pmg_job_pay")
	pmg_job_support = rs("pmg_job_support")
	pmg_jisa_pay = rs("pmg_jisa_pay")
	pmg_long_pay = rs("pmg_long_pay")
	pmg_disabled_pay = rs("pmg_disabled_pay")
	
	pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay
				
	meals_pay = pmg_meals_pay
	car_pay = pmg_car_pay
	meals_tax_pay = 0
	car_tax_pay = 0
	if  meals_pay > 100000 then
	         meals_tax_pay = meals_pay - 100000
			 meals_pay =  100000
	end if
	if car_pay > 200000 then
	         car_tax_pay = car_pay - 200000
			 car_pay =  200000
	end if
	
	pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	pmg_tax_no = meals_pay + car_pay
	
    Sql = "select * from emp_master_month where (emp_month = '"+pmg_yymm+"' ) and (emp_no = '"+emp_no+"')"
	Set Rs_emp = DbConn.Execute(SQL)
	if not Rs_emp.EOF or not Rs_emp.BOF then
	        
			i = i + 1
			
			emp_grade = rs_emp("emp_grade")
			emp_in_date = rs_emp("emp_in_date")
		    emp_position = rs_emp("emp_position")
		    emp_company = rs_emp("emp_company")
			emp_bonbu = rs_emp("emp_bonbu")
			emp_saupbu = rs_emp("emp_saupbu")
			emp_team = rs_emp("emp_team")
			emp_org_code = rs_emp("emp_org_code")
			emp_org_name = rs_emp("emp_org_name")
			emp_reside_place = rs_emp("emp_reside_place") 
			emp_reside_company = rs_emp("emp_reside_company")
			cost_center = rs_emp("cost_center")
			cost_group = rs_emp("cost_group")
'			mg_saupbu = rs_emp("mg_saupbu")
			
			Sql = "SELECT * FROM pay_bank_account where emp_no = '"+emp_no+"'"
            Set rs_bnk = DbConn.Execute(SQL)
            if not rs_bnk.eof then
                  bank_name = rs_bnk("bank_name")
                  account_no = rs_bnk("account_no")
		          account_holder = rs_bnk("account_holder")
	           else
                  bank_name = ""
	    	      account_no = ""
		          account_holder = ""
            end if
            rs_bnk.close()	 
		   
	        'sql = "update pay_month_give set pmg_in_date='"&emp_in_date&"' where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
			
			sql = "update pay_month_give set pmg_date='"&pmg_date&"',pmg_in_date='"&emp_in_date&"',pmg_grade='"&emp_grade&"',pmg_position='"&emp_position&"',pmg_bonbu='"&emp_bonbu&"',pmg_saupbu='"&emp_saupbu&"',pmg_team='"&emp_team&"',pmg_org_name='"&emp_org_name&"',pmg_org_code='"&emp_org_code&"',pmg_reside_place='"&emp_reside_place&"',pmg_reside_company='"&emp_reside_company&"',cost_center='"&cost_center&"',cost_group='"&cost_group&"',pmg_bank_name='"&bank_name&"',pmg_account_no='"&account_no&"',pmg_account_holder='"&account_holder&"',pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_give_total='"&pmg_give_total&"' where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
		
		   dbconn.execute(sql)	 
		   
		   sql = "update pay_month_deduct set de_date='"&pmg_date&"',de_grade='"&emp_grade&"',de_position='"&emp_position&"',de_bonbu='"&emp_bonbu&"',de_saupbu='"&emp_saupbu&"',de_team='"&emp_team&"',de_org_name='"&emp_org_name&"',de_org_code='"&emp_org_code&"',de_reside_place='"&emp_reside_place&"',de_reside_company='"&emp_reside_company&"',cost_group='"&cost_group&"',cost_center='"&cost_center&"' where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '"+pmg_id+"') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
		
		   dbconn.execute(sql)
		   
		else
		   response.write(emp_no)
		   response.write("/")
		   
	end if	 
	    Rs_emp.close()	
	    Rs.MoveNext()
  loop		
		response.write"<script language=javascript>"
		response.write"alert('급여에 조직 데이터가 만들어 졌습니다..."&j&" - "&i&"');"		
		response.write"location.replace('insa_person_mg.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert(' 처리된 내역이없습니다...');"		
		response.write"location.replace('insa_person_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
