<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = "케이원정보통신"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

j = 0

Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') ORDER BY emp_company,emp_no ASC"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    emp_grade = Rs("emp_grade")
	
    if emp_grade = "회장" or emp_grade = "부회장" or emp_grade = "사장" or emp_grade = "부사장" or emp_grade = "총괄대표" then 
	        stock_end_date = "1900-01-01"
	   else
	        emp_no = Rs("emp_no")
			emp_name = Rs("emp_name")
		    emp_company = Rs("emp_company")
			emp_bonbu = Rs("emp_bonbu")
			emp_saupbu = Rs("emp_saupbu")
			emp_team = Rs("emp_team")
			emp_in_date = Rs("emp_in_date")
			emp_org_code = Rs("emp_org_code")
			emp_org_name = Rs("emp_org_name")
			emp_reside_place = Rs("emp_reside_place")
			emp_reside_company = Rs("emp_reside_company")
			org_level = "개인"
			stock_end_date = "1900-01-01"
			
			Sql = "select * from met_stock_code where stock_code = '"+emp_no+"'"
	        Set Rs_stock = DbConn.Execute(SQL)
	        if  Rs_stock.EOF or Rs_stock.BOF then
			
		        j = j + 1
			
	            sql = "insert into met_stock_code (stock_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code,stock_manager_name"
		        sql = sql + ",reg_date,reg_user) values "
		        sql = sql + " ('"&emp_no&"','"&org_level&"','"&emp_name&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_in_date&"','"&stock_end_date&"','"&emp_no&"','"&emp_name&"',now(),'"&user_name&"')"        
			
			    dbconn.execute(sql)	
		    end if
 
	end if	 
'	    Rs_emp.close()	
	    Rs.MoveNext()
  loop		
		response.write"<script language=javascript>"
		response.write"alert('창고 개인 데이터가 만들어 졌습니다......"&j&"');"		
		response.write"location.replace('met_goods_code_mg.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('등록된 조직 내역이없습니다...');"		
		response.write"location.replace('met_goods_code_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
