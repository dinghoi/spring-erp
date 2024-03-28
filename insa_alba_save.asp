<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	draft_no = request.form("draft_no")
	
    draft_man = request.form("draft_man")
    draft_date = request.form("draft_date")
    draft_live_id = ""
    draft_live_name = request.form("draft_live_name")
	if draft_live_name = "거주" then
	       draft_live_id = "1"
	   else
	       draft_live_id = "2"
	end if
    draft_tax_id = request.form("draft_tax_id")
    company = request.form("company")
    bonbu = request.form("bonbu")
    saupbu = request.form("saupbu")
    team = request.form("team")
    org_name = request.form("org_name")
    'cost_company = request.form("cost_company")
	cost_company = ""
    sign_no = request.form("sign_no")
    deposit_date = request.form("deposit_date")
	if deposit_date = "" then
	   deposit_date = "0000-00-00"
	end if
    deposit_man = request.form("deposit_man")
    work_memo = request.form("work_memo")
    bank_name = request.form("bank_name")
    account_no = request.form("account_no")
    account_name = request.form("account_name")
    person_no1 = request.form("person_no1")
    person_no2 = request.form("person_no2")
    nation_id = ""
    nation_name = request.form("nation_name")
    tel_ddd = request.form("tel_ddd")
    tel_no1 = request.form("tel_no1")
    tel_no2 = request.form("tel_no2")
    hp_ddd = request.form("hp_ddd")
    hp_no1 = request.form("hp_no1")
    hp_no2 = request.form("hp_no2")
    e_mail = request.form("e_mail")
    end_yn = request.form("end_yn")
	zip_code = request.form("zip_code")
    sido = request.form("sido")
    gugun = request.form("gugun")
    dong = request.form("dong")
    addr = request.form("addr")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_emp = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect
	
	Sql="select * from emp_etc_code where emp_etc_type = '50' and emp_etc_name = '"&bank_name&"'"
	Rs_etc.Open Sql, Dbconn, 1
	bank_code = rs_etc("emp_etc_code")
	rs_etc.close()

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_alba_mst set draft_man='"&draft_man&"',draft_live_name='"&draft_live_name&"',draft_tax_id='"&draft_tax_id&"',company='"&company&"',org_name='"&org_name&"',cost_company='"&cost_company&"',sign_no='"&sign_no&"',deposit_date='"&deposit_date&"',deposit_man='"&deposit_man&"',work_memo='"&work_memo&"',bank_code='"&bank_code&"',bank_name='"&bank_name&"',account_no='"&account_no&"',account_name='"&account_name&"',e_mail='"&e_mail&"',zip_code='"&zip_code&"',sido='"&sido&"',gugun='"&gugun&"',dong='"&dong&"',addr='"&addr&"',end_yn='"&end_yn&"',mod_date= now(),mod_id='"&emp_user&"' where draft_no ='"&draft_no&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into emp_alba_mst (draft_no,draft_man,draft_date,draft_live_id,draft_live_name,draft_tax_id,company,bonbu,saupbu,team,org_name,cost_company,sign_no,deposit_date,deposit_man,work_memo,bank_code,bank_name,account_no,account_name,person_no1,person_no2,nation_id,nation_name,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,e_mail,zip_code,sido,gugun,dong,addr,end_yn,reg_id,reg_date) values "
		sql = sql +	" ('"&draft_no&"','"&draft_man&"','"&draft_date&"','"&draft_live_id&"','"&draft_live_name&"','"&draft_tax_id&"','"&company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&cost_company&"','"&sign_no&"','"&deposit_date&"','"&deposit_man&"','"&work_memo&"','"&bank_code&"','"&bank_name&"','"&account_no&"','"&account_name&"','"&person_no1&"','"&person_no2&"','"&nation_id&"','"&nation_name&"','"&tel_ddd&"','"&tel_no1&"','"&tel_no2&"','"&hp_ddd&"','"&hp_no1&"','"&hp_no2&"','"&e_mail&"','"&zip_code&"','"&sido&"','"&gugun&"','"&dong&"','"&addr&"','"&end_yn&"','"&emp_user&"',now())"
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
