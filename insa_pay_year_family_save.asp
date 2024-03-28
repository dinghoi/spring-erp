<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	emp_no = request.form("emp_no")
	emp_name = request.form("emp_name")

    incom_year = request.form("incom_year")
	incom_in_date = request.form("incom_in_date")
	incom_grade = request.form("incom_grade")
	incom_emp_type = request.form("incom_emp_type")
	incom_pay_type = request.form("incom_pay_type")
	incom_company = request.form("incom_company")
	incom_org_code = request.form("incom_org_code")
	incom_org_name = request.form("incom_org_name")

    incom_wife_yn = request.form("incom_wife_yn")
    incom_age20 = int(request.form("incom_age20"))
    incom_age60 = int(request.form("incom_age60"))
    incom_old = int(request.form("incom_old"))
    incom_disab = int(request.form("incom_disab"))
    incom_woman = request.form("incom_woman")
	if incom_wife_yn = "1" then 
	      incom_family_cnt = incom_age20 + incom_age60 + incom_old + incom_disab + 1
	   else 
          incom_family_cnt = incom_age20 + incom_age60 + incom_old + incom_disab
    end if
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_emp = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	
	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update pay_year_income set incom_family_cnt='"&incom_family_cnt&"',incom_wife_yn='"&incom_wife_yn&"',incom_age20='"&incom_age20&"',incom_age60='"&incom_age60&"',incom_old='"&incom_old&"',incom_disab='"&incom_disab&"',incom_woman='"&incom_woman&"',incom_mod_date= now(),incom_mod_user='"&emp_user&"' where incom_emp_no = '"&emp_no&"' and incom_year = '"&incom_year&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into pay_year_income (incom_emp_no,incom_year,incom_emp_name,incom_in_date,incom_grade,incom_emp_type,incom_pay_type,incom_company,incom_org_code,incom_org_name,incom_family_cnt,incom_wife_yn,incom_age20,incom_age60,incom_old,incom_disab,incom_woman,incom_reg_date,incom_reg_user) values "
		sql = sql +	" ('"&emp_no&"','"&incom_year&"','"&emp_name&"','"&incom_in_date&"','"&incom_grade&"','"&incom_emp_type&"','"&incom_pay_type&"','"&incom_company&"','"&incom_org_code&"','"&incom_org_name&"','"&incom_family_cnt&"','"&incom_wife_yn&"','"&incom_age20&"','"&incom_age60&"','"&incom_old&"','"&incom_disab&"','"&incom_woman&"',now(),'"&emp_user&"')"
		
		dbconn.execute(sql)
		
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
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
