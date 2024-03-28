<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

' 조직신설이 되면 다른 계열사에도 추가로 등록 시키는

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = now()

be_org_code = "1260"
net_org_code = "4260"
ko_org_code = "6260"
'response.write(be_year)
'response.write("/")
'response.write(af_year)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Dbconn.BeginTrans 


' 조직 체크
Sql = "SELECT * FROM emp_org_mst where org_code = '"&be_org_code&"'"
Set rs=DbConn.Execute(Sql)

    org_level = rs("org_level")
    org_name = rs("org_name")
    org_date = rs("org_date")
	org_end_date = rs("org_end_date")
    org_empno = rs("org_empno")
    org_empname = rs("org_emp_name")
    org_company = rs("org_company")
    org_bonbu = rs("org_bonbu")
    org_saupbu = rs("org_saupbu")
    org_team = rs("org_team")
	org_reside_place = rs("org_reside_place")
	org_reside_company = rs("org_reside_company")
	org_cost_group = rs("org_cost_group")
	org_cost_center = rs("org_cost_center")
    owner_org = rs("org_owner_org")
    owner_empno = rs("org_owner_empno")
    owner_empname = rs("org_owner_empname")
	if rs("org_table_org") = "" or isnull(rs("org_table_org")) then
	        org_table_org = 0
	   else	
			org_table_org = rs("org_table_org")
	end if
    tel_ddd = rs("org_tel_ddd")
    tel_no1 = rs("org_tel_no1")
    tel_no2 = rs("org_tel_no2")
	org_sido = rs("org_sido")
    org_gugun = rs("org_gugun")
    org_dong = rs("org_dong")
    org_addr = rs("org_addr")
    org_end_date = rs("org_end_date")
    org_reg_date = rs("org_reg_date")
	org_reg_user = rs("org_reg_user")
    org_mod_date = rs("org_mod_date")
    org_mod_user = rs("org_mod_user")
	   
'	    sql = "insert into emp_org_mst (org_code,org_level,org_company,org_bonbu,org_saupbu,org_team,org_name,org_reside_place,org_reside_company,org_cost_group,org_empno,org_emp_name,org_date,org_tel_ddd,org_tel_no1,org_tel_no2,org_cost_center"
        sql = sql + ",org_owner_org,org_owner_empno,org_owner_empname,org_table_org,org_sido,org_gugun,org_dong,org_addr"
		sql = sql + ",org_reg_date,org_reg_user) values "
		sql = sql + " ('"&net_org_code&"','"&org_level&"','케이네트웍스','"&org_bonbu&"','"&org_saupbu&"','"&org_team&"','"&org_name&"','"&org_reside_place&"','"&org_reside_company&"','"&org_cost_group&"','"&org_empno&"','"&org_empname&"','"&org_date&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&org_cost_center&"','"&owner_org&"','"&owner_empno&"','"&owner_empname&"','"&org_table_org&"','"&org_sido&"','"&org_gugun&"','"&org_dong&"','"&org_addr&"',now(),'"&org_reg_user&"')"
	   
'	   dbconn.execute(sql)
	   
	   sql = "insert into emp_org_mst (org_code,org_level,org_company,org_bonbu,org_saupbu,org_team,org_name,org_reside_place,org_reside_company,org_cost_group,org_empno,org_emp_name,org_date,org_tel_ddd,org_tel_no1,org_tel_no2,org_cost_center"
        sql = sql + ",org_owner_org,org_owner_empno,org_owner_empname,org_table_org,org_sido,org_gugun,org_dong,org_addr"
		sql = sql + ",org_reg_date,org_reg_user) values "
		sql = sql + " ('"&ko_org_code&"','"&org_level&"','코리아디엔씨','"&org_bonbu&"','"&org_saupbu&"','"&org_team&"','"&org_name&"','"&org_reside_place&"','"&org_reside_company&"','"&org_cost_group&"','"&org_empno&"','"&org_empname&"','"&org_date&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&org_cost_center&"','"&owner_org&"','"&owner_empno&"','"&owner_empname&"','"&org_table_org&"','"&org_sido&"','"&org_gugun&"','"&org_dong&"','"&org_addr&"',now(),'"&org_reg_user&"')"
	   
	   dbconn.execute(sql)

	
if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	response.write"<script language=javascript>"
	response.write"alert('"&be_org_code&"...신규조직 타 계열사가 추가 되었습니다...');"		
	'response.write"location.replace('insa_master_month_mg.asp');"
	response.write"location.replace('insa_person_mg.asp');"
	response.write"</script>"
	Response.End
end if

dbconn.Close()
Set dbconn = Nothing
	
%>
