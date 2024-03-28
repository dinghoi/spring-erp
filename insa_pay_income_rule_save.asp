<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	rule_id = request.form("rule_id")
	rule_cl = request.form("rule_cl")
	rule_id_name = request.form("rule_id_name")

	if	u_type = "U" then
	       rule_yyyy = request.form("rule_yyyy")
		else
	       from_date = request.form("from_date")
	       rule_yyyy = mid(cstr(from_date),1,4)
    end if
	
	rule_year_pay = int(request.form("rule_year_pay"))
	rule_st_deduct = int(request.form("rule_st_deduct"))
	rule_exceed = int(request.form("rule_exceed"))
	rule_add = int(request.form("rule_add"))
	
	rule_exceed_rate = request.form("rule_exceed_rate")
	rule_add_rate = request.form("rule_add_rate")
	rule_comment = request.form("rule_comment")
	
	'start_time = cstr(start_hh) + cstr(start_mm)
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "Update pay_income_rule set rule_year_pay='"&rule_year_pay&"',rule_st_deduct ='"&rule_st_deduct&"',rule_exceed ='"&rule_exceed&"',rule_add='"&rule_add&"',rule_exceed_rate='"&rule_exceed_rate&"',rule_add_rate='"&rule_add_rate&"',rule_comment='"&rule_comment&"' where rule_yyyy = '"&rule_yyyy&"' and rule_id = '"&rule_id&"' and rule_cl = '"&rule_cl&"'"
		dbconn.execute(sql)
		
	  else
		sql="insert into pay_income_rule (rule_yyyy,rule_id,rule_cl,rule_id_name,rule_year_pay,rule_st_deduct,rule_exceed,rule_add,rule_exceed_rate,rule_add_rate,rule_comment,rule_reg_user,rule_reg_date) values ('"&rule_yyyy&"','"&rule_id&"','"&rule_cl&"','"&rule_id_name&"','"&rule_year_pay&"','"&rule_st_deduct&"','"&rule_exceed&"','"&rule_add&"','"&rule_exceed_rate&"','"&rule_add_rate&"','"&rule_comment&"','"&emp_user&"',now())"
		dbconn.execute(sql)

	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "자장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
