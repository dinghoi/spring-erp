<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next
	
    curr_date = mid(cstr(now()),1,10)
	
	u_type = request.form("u_type")
	
	reg_user = request.cookies("nkpmg_user")("coo_user_name")
	mod_user = request.cookies("nkpmg_user")("coo_user_name")
	
	org_level = request.form("org_level")
	org_code = request.form("org_code")
	org_name = request.form("org_name")
	org_date = request.form("org_date")
	org_empno = request.form("org_empno")
	org_empname = request.form("org_empname")
	org_company = request.form("org_company")
	org_bonbu = request.form("org_bonbu")
	org_saupbu = request.form("org_saupbu")
	org_team = request.form("org_team")
	
	org_cost_group = request.form("org_cost_group")
	org_cost_center = request.form("org_cost_center")
	
	if org_bonbu = "" or isnull(org_bonbu) then
	       org_bonbu = ""
    end if
	if org_saupbu = "" or isnull(org_saupbu) then
	       org_saupbu = ""
    end if
	if org_team = "" or isnull(org_team) then
	       org_team = ""
    end if
	if org_level = "회사" then
	      org_company = org_name
	   elseif org_level = "본부" then
	              org_bonbu = org_name
			  elseif org_level = "사업부" then
	                     org_saupbu = org_name
					 elseif org_level = "팀" then
	                           org_team = org_name
	end if
	org_reside_company = request.form("org_reside_company")
	if org_reside_company = "" or isnull(org_reside_company) then
	       org_reside_company = ""
    end if

	if org_cost_group = "" or isnull(org_cost_group) then
	       org_cost_group = org_reside_company
	end if
	
	owner_org = request.form("owner_org")
	owner_orgname = request.form("owner_orgname")
	owner_empno = request.form("owner_empno")
	owner_empname = request.form("owner_empname")
	org_table_org = int(request.form("org_table_org"))
	org_zip = request.form("org_zip")
	org_sido = request.form("org_sido")
	org_gugun = request.form("org_gugun")
	org_dong = request.form("org_dong")
	org_addr = request.form("org_addr")
	org_end_date = request.form("org_end_date")
    tel_ddd = request.form("tel_ddd")
	tel_no1 = request.form("tel_no1") 
	tel_no2 = request.form("tel_no2")
    if tel_ddd = "" then
	   tel_ddd = ""
	   tel_no1 = ""
	   tel_no2 = ""
	end if
	org_reside_place = request.form("org_reside_place")
	if org_level = "상주처" then
			  org_cost_center = "상주직접비"
	   else
	          org_cost_group = org_saupbu
'			  org_reside_company = ""
			  if org_saupbu = "" then
			       if org_bonbu = ""  then
				          org_cost_group = org_company
					  else  
						  org_cost_group = org_bonbu
				   end if
			  end if
	end if
	
	if org_reside_company <> "" then
	   org_cost_group = request.form("org_cost_group")
	end if
	
	if isnull(org_date) or org_date = "" then
	    org_date = "0000-00-00"
	end if
	if isnull(org_end_date) or org_end_date = "" then
	    org_end_date = "0000-00-00"
	end if
	
	org_reg_date = request.form("org_reg_date")
	org_mod_date = request.form("org_mod_date")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_stock = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "update emp_org_mst set org_level='"&org_level&"',org_company='"&org_company&"',org_bonbu='"&org_bonbu&"',org_saupbu='"&org_saupbu&"',org_team='"&org_team&"',org_name='"&org_name&"',org_reside_place='"&org_reside_place&"',org_reside_company='"&org_reside_company&"',org_cost_group='"&org_cost_group&"',org_empno='"&org_empno&"',org_emp_name='"&org_empname&"',org_date='"&org_date&"',org_tel_ddd='"&tel_ddd&"',org_tel_no1='"&tel_no1&"',org_tel_no2='"&tel_no2&"',org_owner_org='"&owner_org&"',org_owner_empno='"&owner_empno&"',org_owner_empname='"&owner_empname&"',org_table_org='"&org_table_org&"',org_sido='"&org_sido&"',org_gugun='"&org_gugun&"',org_dong='"&org_dong&"',org_addr='"&org_addr&"',org_cost_group='"&org_cost_group&"',org_cost_center='"&org_cost_center&"',org_end_date='"&org_end_date&"',org_mod_date=now(),org_mod_user='"&mod_user&"' where org_code = '"&org_code&"'"

		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into emp_org_mst (org_code,org_level,org_company,org_bonbu,org_saupbu,org_team,org_name,org_reside_place,org_reside_company,org_cost_group,org_empno,org_emp_name,org_date,org_tel_ddd,org_tel_no1,org_tel_no2,org_cost_center"
        sql = sql + ",org_owner_org,org_owner_empno,org_owner_empname,org_table_org,org_sido,org_gugun,org_dong,org_addr"
		sql = sql + ",org_reg_date,org_reg_user) values "
		sql = sql + " ('"&org_code&"','"&org_level&"','"&org_company&"','"&org_bonbu&"','"&org_saupbu&"','"&org_team&"','"&org_name&"','"&org_reside_place&"','"&org_reside_company&"','"&org_cost_group&"','"&org_empno&"','"&org_empname&"','"&org_date&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&org_cost_center&"','"&owner_org&"','"&owner_empno&"','"&owner_empname&"','"&org_table_org&"','"&org_sido&"','"&org_gugun&"','"&org_dong&"','"&org_addr&"',now(),'"&reg_user&"')"
		
		'response.write sql
		
		dbconn.execute(sql)
		
'		if org_level = "회사" or org_level = "본부" or org_level = "사업부" or org_level = "팀" then 
'		    sql = "insert into met_stock_code (stock_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_open_date,stock_manager_code,stock_manager_name"
'		    sql = sql + ",reg_date,reg_user) values "
'		    sql = sql + " ('"&org_code&"','"&org_level&"','"&org_name&"','"&org_company&"','"&org_bonbu&"','"&org_saupbu&"','"&org_team&"','"&org_date&"','"&org_empno&"','"&org_emp_name&"',now(),'"&user_name&"')"        
			
'			dbconn.execute(sql)	
'		end if
		
	end if
' 창고코드 등록	 
if org_level = "본사" or org_level = "팀" then 
  if org_code <> "" or org_code <> " " then
    sql="select * from met_stock_code where stock_code='"&org_code&"'"
	set rs_stock=dbconn.execute(sql)

    if rs_stock.eof then
       stock_end_date = "1900-01-01"
	   if org_level = "회사" then 
	          stock_level = "본사"
		  else
	          stock_level = "팀"
	   end if
	   sql = "insert into met_stock_code (stock_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code,stock_manager_name"
		        sql = sql + ",reg_date,reg_user) values "
		        sql = sql + " ('"&org_code&"','"&stock_level&"','"&org_name&"','"&org_company&"','"&org_bonbu&"','"&org_saupbu&"','"&org_team&"','"&org_date&"','"&stock_end_date&"','"&org_empno&"','"&org_empname&"',now(),'"&reg_user&"')"        

		'response.write(sql)
		dbconn.execute(sql)	 
	else
	    sql = "update met_stock_code set stock_name='"&org_name&"',stock_company='"&org_company&"',stock_bonbu='"&org_bonbu&"',stock_saupbu='"&org_saupbu&"',stock_team='"&org_team&"',stock_open_date='"&org_date&"',stock_manager_code='"&org_empno&"',stock_manager_name='"&org_empname&"' where stock_code='"&org_code&"'"

		'response.write sql
		
		dbconn.execute(sql)	  
    end if
  end if
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
	'response.write"location.replace('insa_org.asp');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"			
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
