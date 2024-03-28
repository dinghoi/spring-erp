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

view_condi=Request.form("view_condi1")
view_c=Request.form("view_c1")
field_check=Request.form("field_check1")
field_bonbu=Request.form("field_bonbu1")
field_saupbu=Request.form("field_saupbu1")
field_team=Request.form("field_team1")
field_org=Request.form("field_org1")

'response.write(view_condi)
'response.write(view_c)
'response.write(field_check)
'response.write(field_bonbu)
'response.write(field_saupbu)
'response.write(field_team)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If view_c = "" Then
	field_check = "total"
	view_c = "bonbu"
End If

If field_check = "total" Then
       owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"')"
	   field_check = ""
   else
       if view_c = "bonbu" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_bonbu like '%" + field_bonbu + "%')"
       end if
	   if view_c = "saupbu" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_saupbu like '%" + field_saupbu + "%')"
       end if
	   if view_c = "team" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_team like '%" + field_team + "%')"
       end if
	   if view_c = "orgm" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_org_name like '%" + field_org + "%')"
       end if
End If

i = 0
j = 0

sql = "select * from emp_master " + owner_sql 
Rs.Open Sql, Dbconn, 1

if not Rs.eof then
   do until Rs.eof
          emp_no = rs("emp_no")
		  emp_company = rs("emp_company")
		  emp_org_code = rs("emp_org_code")
		  
		  j = j + 1

          sql = "select * from emp_org_mst where (org_code = '"&emp_org_code&"')"
		  Set Rs_org = DbConn.Execute(SQL)
	      if not Rs_org.eof then	
		         org_company = Rs_org("org_company")
		         org_bonbu = Rs_org("org_bonbu")
		         org_saupbu = Rs_org("org_saupbu")
		         org_team = Rs_org("org_team")
		         org_name = Rs_org("org_name")
				 
				 i = i + 1
				 
				 sql = "update emp_master set emp_company='"&org_company&"',emp_bonbu='"&org_bonbu&"',emp_saupbu='"&org_saupbu&"',emp_team='"&org_team&"',emp_mod_date=now(),emp_mod_user='"&mod_user&"' where emp_no = '"&emp_no&"'"
				 
				 dbconn.execute(sql)
		
		         sql="select * from memb where user_id='"&emp_no&"'"
	             set rs_memb=dbconn.execute(sql)
                 if not rs_memb.eof then
		             sql = "update memb set emp_company='"&org_company&"',bonbu='"&org_bonbu&"',saupbu='"&org_saupbu&"',team='"&org_team&"',mod_date=now() where user_id='"&emp_no&"'"
			
			         dbconn.execute(sql)
		          end if
	       end if
		Rs.MoveNext()
    loop		
		response.write"<script language=javascript>"
		response.write"alert('해당 조직 직원의 상위 조직이 변경 되었습니다..."&j&" --> "&i&"');"		
		response.write"location.replace('insa_emp_owner_org_list.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('처리할 내역이 없습니다...');"		
		response.write"location.replace('insa_emp_owner_org_list.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
