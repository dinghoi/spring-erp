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

rever_yyyy=Request.form("rever_yyyy")
target_date=Request.form("target_date")
view_condi=Request.form("view_condi")

'response.write(rever_yyyy)
'response.write(view_condi)
'response.write(target_date)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_no,emp_name ASC"
Rs.Open Sql, Dbconn, 1

'Sql = "SELECT * FROM emp_year_leave WHERE year_year='" + rever_yyyy + "' and year_empno='" +rs("emp_no") + "'"	
Sql = "SELECT * FROM emp_year_leave WHERE year_year='" + rever_yyyy + "'"	
Set Rs_year=Dbconn.Execute(sql)
if Rs_year.eof then
   do until rs.eof

       year_empno = rs("emp_no")
	   year_year = rever_yyyy
	   year_emp_name = rs("emp_name")
	   year_in_date = rs("emp_in_date")
	   year_first_date = rs("emp_first_date")
	   year_yuncha_date = rs("emp_yuncha_date")
	   year_company = rs("emp_company")
	   year_bonbu = rs("emp_bonbu")
	   year_saupbu = rs("emp_saupbu")
	   year_team = rs("emp_team")
	   year_org_code = rs("emp_org_code")
	   year_org_name = rs("emp_org_name")
	   
	   year_use_count = 0
	   year_remain_count = 0
	   
       if rs("emp_yuncha_date") = "1900-01-01" or isNull(rs("emp_yuncha_date")) then
            emp_yuncha_date = rs("emp_in_date")
         else 
            emp_yuncha_date = rs("emp_yuncha_date")
       end if

      ' 근속년수
	  'target_date1 = from_date + 1
      year_cnt = datediff("yyyy", emp_yuncha_date, target_date)
						  							  
	  ' 연차일수
	  if (datediff("d", emp_yuncha_date, target_date) + 1) / 365 < 1 then
             yun_day = datediff("m", emp_yuncha_date, target_date) 
		 else
		     yun_day = round((((datediff("d", emp_yuncha_date, target_date) + 1) / 365) / 2),0) + 14
	  end if
							  
	  ' 누적연차수
	  if datediff("yyyy", emp_yuncha_date, target_date) mod 2 = 1 then
	          tot_yun = round(((year_cnt ^ 2 + 58 * year_cnt - 0) / 4),0)
		 else
	          tot_yun = year_cnt / 2 * (year_cnt / 2 + 1) + 14 * year_cnt
	  end if
						  
      mon_cnt = datediff("m", emp_yuncha_date, target_date) 
						  
	  if mon_cnt < 0 then
	        mon_cnt = 0
			yun_day = 0
	  end if

      year_continu_year = year_cnt
	  year_continu_month = mon_cnt
	  year_basic_count = yun_day
	  year_add_count = 0
	  year_leave_count = tot_yun

	   sql="insert into emp_year_leave (year_empno,year_year,year_emp_name,year_first_date,year_in_date,year_yuncha_date,year_company,year_bonbu,year_saupbu,year_team,year_org_code,year_org_name,year_continu_year,year_continu_month,year_basic_count,year_add_count,year_leave_count,year_use_count,year_remain_count,year_reg_date,year_reg_user) values ('"&year_empno&"','"&year_year&"','"&year_emp_name&"','"&year_first_date&"','"&year_in_date&"','"&year_yuncha_date&"','"&year_company&"','"&year_bonbu&"','"&year_saupbu&"','"&year_team&"','"&year_org_code&"','"&year_org_name&"','"&year_continu_year&"','"&year_continu_month&"','"&year_basic_count&"','"&year_add_count&"','"&year_leave_count&"','"&year_use_count&"','"&year_remain_count&"',now(),'"&emp_user&"')"
  
	   dbconn.execute(sql)
	   
	   Rs.MoveNext()
    loop		
		response.write"<script language=javascript>"
		response.write"alert('연차휴가일수 데이터가 만들어 졌습니다...');"		
		response.write"location.replace('insa_year_leave_bat.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('이미 처리된 내역이 있습니다...');"		
		response.write"location.replace('insa_year_leave_bat.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
