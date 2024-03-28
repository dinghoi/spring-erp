<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	u_type = request.form("u_type")
	insure_year = request.form("insure_year")
	old_insure_year = request.form("old_insure_year")
	nps_per = request.form("nps_per")
	nhis_per = request.form("nhis_per")
	longcare_per = request.form("longcare_per")
	epi_person_per = request.form("epi_person_per")
	epi_company_per = request.form("epi_company_per")
	comwel_per = request.form("comwel_per")
	insure_tot_per = request.form("insure_tot_per")
	income_tax_per = request.form("income_tax_per")
	annual_pay_per = request.form("annual_pay_per")
	retire_pay_per = request.form("retire_pay_per")
	person_tot_per = request.form("person_tot_per")
	insure_memo = request.form("insure_memo")

	if u_type = "U" then
		sql = "delete from insure_per where insure_year ='"&old_insure_year&"'"
		dbconn.execute(sql)
	end if


	sql="insert into insure_per  values ('"&insure_year&"','"&nps_per&"','"&nhis_per&"','"&longcare_per&"','"&epi_person_per&"','"&epi_company_per&"',"&comwel_per&",'"&insure_tot_per&"','"&income_tax_per&"',"&annual_pay_per&",'"&retire_pay_per&"','"&person_tot_per&"','"&insure_memo&"','"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.Redirect "insure_per_mg.asp"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
