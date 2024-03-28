<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")

	draft_no = request.form("draft_no")
	rever_yymm = request.form("rever_yymm")
	give_date = request.form("give_date")
	old_date = request.form("old_date")
	draft_man = request.form("draft_man")
	draft_tax_id = request.form("draft_tax_id")
	company = request.form("emp_company")
	bonbu = request.form("bonbu")
	saupbu = request.form("saupbu")
	team = request.form("team")
	cost_company = request.form("cost_company")
	org_name = request.form("org_name")
	sign_no = request.form("sign_no")
	work_comment = request.form("work_comment")
	bank_name = request.form("bank_name")
	account_no = request.form("account_no")
	account_name = request.form("account_name")
	
	alba_pay =int(request.form("alba_pay"))
	alba_trans = int(request.form("alba_trans"))
	alba_meals = int(request.form("alba_meals"))
	alba_other = int(request.form("alba_other"))
	alba_other2 = 0
	alba_give_total = int(request.form("give_tot"))
	curr_pay = int(request.form("curr_pay"))
	de_other = int(request.form("de_other"))
	tax_amt1 = int(request.form("tax_amt1"))
	tax_amt2 = int(request.form("tax_amt2"))
	alba_cnt = int(request.form("alba_cnt"))
	alba_work = int(request.form("alba_work"))
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	emp_user = user_name

	if	u_type = "U" then
		sql = "delete from pay_alba_cost where rever_yymm = '"&rever_yymm&"' and draft_no = '"&draft_no&"' and give_date = '"&old_date&"'"
		dbconn.execute(sql)
	end if
'		sql = "Update pay_alba_cost set alba_pay='"&alba_pay&"',alba_trans ='"&alba_trans&"',alba_meals ='"&alba_meals&"',alba_other='"&alba_other&"',alba_other2='"&alba_other2&"',alba_give_total='"&alba_give_total&"',tax_amt1='"&tax_amt1&"',tax_amt2='"&tax_amt2&"',pay_amount='"&curr_pay&"',alba_cnt='"&alba_cnt&"',alba_work='"&alba_work&"',work_comment='"&work_comment&"',bank_name='"&bank_name&"',account_no='"&account_no&"',account_name='"&account_name&"',mod_id='"&emp_user&"',mod_date=now() where rever_yymm = '"&rever_yymm&"' and draft_no = '"&draft_no&"'"
'		dbconn.execute(sql)
'	  else
	sql="insert into pay_alba_cost (rever_yymm,draft_no,company,draft_man,draft_tax_id,give_date,bonbu,saupbu,team,org_name,cost_company,sign_no,alba_cnt,alba_work,work_comment,alba_pay,alba_trans,alba_meals,alba_other,alba_other2,alba_give_total,tax_amt1,tax_amt2,de_other,pay_amount,bank_name,account_no,account_name,reg_id,reg_date) values ('"&rever_yymm&"','"&draft_no&"','"&company&"','"&draft_man&"','"&draft_tax_id&"','"&give_date&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&cost_company&"','"&sign_no&"','"&alba_cnt&"','"&alba_work&"','"&work_comment&"','"&alba_pay&"','"&alba_trans&"','"&alba_meals&"','"&alba_other&"','"&alba_other2&"','"&alba_give_total&"','"&tax_amt1&"','"&tax_amt2&"','"&de_other&"','"&curr_pay&"','"&bank_name&"','"&account_no&"','"&account_name&"','"&emp_user&"',now())"
		dbconn.execute(sql)

'	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"opener.document.frm.submit();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
