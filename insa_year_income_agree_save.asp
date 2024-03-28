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

from_date = curr_year + "-" + "01" + "-" + "01"
to_date = curr_year + "-" + "12" + "-" + "31"

emp_no=Request("emp_no")
emp_name=Request("emp_name")
agree_year=Request("agree_year")
agree_year=curr_year

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_agree = Server.CreateObject("ADODB.Recordset")
Set rs_max = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where  (emp_no = '"+emp_no+"')"
Rs.Open Sql, Dbconn, 1

Sql = "SELECT * FROM emp_agree WHERE (agree_year = '" + agree_year + "') and (agree_empno = '" + emp_no + "') and (agree_seq = '001')"	
response.write(sql)
Set Rs_year=Dbconn.Execute(sql)
if Rs_year.eof then
       agree_empno = rs("emp_no")
	   agree_year = agree_year
	   agree_id = "연봉근로계약서"     
	   agree_empname = rs("emp_name")
	   agree_company = rs("emp_company")
	   agree_org_code = rs("emp_org_code")
	   agree_org_name = rs("emp_org_name")
	   agree_grade = rs("emp_grade")
	   agree_job = rs("emp_job")
	   agree_position = rs("emp_position")
	   agree_jikmu = rs("emp_jikmu")
	   agree_emp_type = "정직"
	   agree_birthday = rs("emp_birthday")
	   agree_in_date = rs("emp_in_date")
	   agree_person1 = rs("emp_person1")
	   agree_person2 = rs("emp_person2")
	   agree_date = curr_date
	   agree_sido = rs("emp_sido")
	   agree_gugun = rs("emp_gugun")
	   agree_dong = rs("emp_dong")
	   agree_addr = rs("emp_addr")
	   agree_tel_ddd = rs("emp_tel_ddd")
	   agree_tel_no1 = rs("emp_tel_no1")
	   agree_tel_no2 = rs("emp_tel_no2")
	   agree_from_date = from_date
	   agree_to_date = to_date
	   
	   Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&curr_year&"'"
       Set rs_year = DbConn.Execute(SQL)
           if not rs_year.eof then
                  incom_base_pay = rs_year("incom_base_pay")
                  incom_overtime_pay = rs_year("incom_overtime_pay")
	              incom_meals_pay = rs_year("incom_meals_pay")
                  incom_severance_pay = rs_year("incom_severance_pay")
	              incom_total_pay = rs_year("incom_total_pay")
				  incom_first3_percent = rs_year("incom_first3_percent")
              else
                  incom_base_pay = 0
                  incom_overtime_pay = 0
                  incom_meals_pay = 0
	              incom_severance_pay = 0
                  incom_total_pay = 0
				  incom_first3_percent = 0
            end if
            rs_year.close()
	   
	   agree_base_pay = incom_base_pay
	   agree_extend_pay = incom_overtime_pay
	   agree_meal_pay = incom_meals_pay
	   agree_severance_pay = incom_severance_pay
	   agree_total_pay = incom_total_pay
	   agree_pay_percent = incom_first3_percent
	   
       sql="select max(agree_seq) as max_seq from emp_agree where agree_empno = '"&emp_no&"' and agree_year = '"&agree_year&"'"
	   set rs_max=dbconn.execute(sql)
	
	   if	isnull(rs_max("max_seq"))  then
		    seq_last = "001"
	      else
		    max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		    seq_last = right(max_seq,3)
	   end if
       rs_max.close()
       agree_seq = seq_last
   
	   sql="insert into emp_agree (agree_empno,agree_year,agree_seq,agree_id,agree_empname,agree_company,agree_org_code,agree_org_name,agree_grade,agree_job,agree_position,agree_jikmu,agree_emp_type,agree_birthday,agree_in_date,agree_person1,agree_person2,agree_date,agree_sido,agree_gugun,agree_dong,agree_addr,agree_tel_ddd,agree_tel_no1,agree_tel_no2,agree_from_date,agree_to_date,agree_base_pay,agree_extend_pay,agree_meal_pay,agree_severance_pay,agree_total_pay,agree_pay_percent,agree_reg_date,agree_reg_user) values ('"&agree_empno&"','"&agree_year&"','"&agree_seq&"','"&agree_id&"','"&agree_empname&"','"&agree_company&"','"&agree_org_code&"','"&agree_org_name&"','"&agree_grade&"','"&agree_job&"','"&agree_position&"','"&agree_jikmu&"','"&agree_emp_type&"','"&agree_birthday&"','"&agree_in_date&"','"&agree_person1&"','"&agree_person2&"','"&agree_date&"','"&agree_sido&"','"&agree_gugun&"','"&agree_dong&"','"&agree_addr&"','"&agree_tel_ddd&"','"&agree_tel_no1&"','"&agree_tel_no2&"','"&agree_from_date&"','"&agree_to_date&"','"&agree_base_pay&"','"&agree_extend_pay&"','"&agree_meal_pay&"','"&agree_severance_pay&"','"&agree_total_pay&"','"&agree_pay_percent&"',now(),'"&emp_user&"')"
  
	   dbconn.execute(sql)
	   
		response.write"<script language=javascript>"
		response.write"alert('연봉근로계약을 동의하셨습니다...');"		
		response.write"location.replace('insa_year_leave_bat.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('이미 연봉근로계약을 하셨습니다...');"		
		response.write"location.replace('insa_individual_agree.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
