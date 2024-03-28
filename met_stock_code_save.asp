<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next
	
    curr_date = mid(cstr(now()),1,10)
	
	u_type = request.form("u_type")
	
	user_name = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")
	
	stock_level = request.form("stock_level")
	stock_code = request.form("stock_code")
	stock_name = request.form("stock_name")
	stock_manager_code = request.form("stock_manager_code")
	stock_manager_name = request.form("stock_manager_name")
	stock_go_man = request.form("stock_go_man")
	stock_go_name = request.form("stock_go_name")
	stock_in_man = request.form("stock_in_man")
	stock_in_name = request.form("stock_in_name")
	stock_company = request.form("stock_company")
	stock_bonbu = request.form("stock_bonbu")
	stock_saupbu = request.form("stock_saupbu")
	stock_team = request.form("stock_team")
	stock_open_date = request.form("stock_open_date")
	stock_end_date = request.form("stock_end_date")
	
	if stock_end_date = "" or isnull(stock_end_date) then
	   stock_end_date = "1900-01-01"
	end if

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "update met_stock_code set stock_level='"&stock_level&"',stock_name='"&stock_name&"',stock_company='"&stock_company&"',stock_bonbu='"&stock_bonbu&"',stock_saupbu='"&stock_saupbu&"',stock_team='"&stock_team&"',stock_open_date='"&stock_open_date&"',stock_end_date='"&stock_end_date&"',stock_manager_code='"&stock_manager_code&"',stock_manager_name='"&stock_manager_name&"',stock_go_man='"&stock_go_man&"',stock_go_name='"&stock_go_name&"',stock_in_man='"&stock_in_man&"',stock_in_name='"&stock_in_name&"',mod_date=now(),mod_user='"&user_name&"' where stock_code = '"&stock_code&"'"

		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into met_stock_code (stock_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code,stock_manager_name,stock_go_man,stock_go_name,stock_in_man,stock_in_name"
		sql = sql + ",reg_date,reg_user) values "
		sql = sql + " ('"&stock_code&"','"&stock_level&"','"&stock_name&"','"&stock_company&"','"&stock_bonbu&"','"&stock_saupbu&"','"&stock_team&"','"&stock_open_date&"','"&stock_end_date&"','"&stock_manager_code&"','"&stock_manager_name&"','"&stock_go_man&"','"&stock_go_name&"','"&stock_in_man&"','"&stock_in_name&"',now(),'"&user_name&"')"
		
		'response.write sql
		
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
