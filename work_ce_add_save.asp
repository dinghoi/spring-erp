<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	Dim mg_ce_id(30)
	Dim saupbu_tab(30)
	Dim team_tab(30)
	Dim reside_place_tab(30)
	Dim reside_company_tab(30)
	Dim bonbu_tab(30)
	Dim com_tab(30)
	Dim reside_tab(30)
	Dim org_name_tab(30)
	for i = 1 to 30
		mg_ce_id(i) = ""
		com_tab(i) = ""
		bonbu_tab(i) = ""
		saupbu_tab(i) = ""
		team_tab(i) = ""
		reside_place_tab(i) = ""
		reside_company_tab(i) = ""
		reside_tab(i) = ""
		org_name_tab(i) = ""
	next
	acpt_no = request.form("acpt_no")
	as_type = request.form("as_type")
	work_man_cnt = int(request.form("work_man_cnt"))
	dev_inst_cnt = int(request.form("dev_inst_cnt"))
	ran_cnt = int(request.form("ran_cnt"))
	alba_cnt = int(request.form("alba_cnt"))
	company = request.form("company")

	person_amt = clng((dev_inst_cnt + ran_cnt) / work_man_cnt)
	first_cnt = (dev_inst_cnt + ran_cnt) - (person_amt * (work_man_cnt - 1))
	if first_cnt < 1 then
		first_cnt = 1
	end if

	i = int(work_man_cnt)
	for j = 1 to i
		mg_ce_id(j) = request.form("mg_ce_id"&j)
		com_tab(j) = request.form("emp_company"&j)
		bonbu_tab(j) = request.form("bonbu"&j)
		saupbu_tab(j) = request.form("saupbu"&j)
		team_tab(j) = request.form("team"&j)
		reside_place_tab(j) = request.form("reside_place"&j)
		reside_company_tab(j) = request.form("reside_company"&j)
		reside_tab(j) = request.form("reside"&j)
		org_name_tab(j) = request.form("org_name"&j)
	next

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "delete from ce_work where acpt_no ="&int(acpt_no)
	dbconn.execute(sql)

	for j = 1 to i
		if j = 1 then
			person_cnt = first_cnt
		  else
			person_cnt = person_amt
		end if
		sql="INSERT INTO ce_work (acpt_no, mg_ce_id, work_id, as_type, company, emp_company, bonbu, saupbu, team, org_name, reside_place, reside"& _
		",reside_company,work_man_cnt"&",dev_inst_cnt,ran_cnt,alba_cnt,person_amt,reg_id,reg_date) values ('"&acpt_no&"','"&mg_ce_id(j)& _
		"','2','"&as_type&"','"&company&"','"&com_tab(j)&"','"&bonbu_tab(j)&"','"&saupbu_tab(j)&"','"&team_tab(j)&"','"&org_name_tab(j)& _
		"','"&reside_place_tab(j)&"','"&reside_tab(j)&"','"&reside_company_tab(j)&"',"&work_man_cnt&","&dev_inst_cnt&","&ran_cnt& _
		","&alba_cnt&","&person_cnt&",'"&user_id&"',now())"
		dbconn.execute(sql)

	next

'	sql = "Update as_acpt set overtime ='Y' where acpt_no ="&int(acpt_no)
'	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "등록되었습니다...."
	end if
'	ran_cnt = 33

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"opener.document.frm.dev_inst_cnt.value = '"&dev_inst_cnt&"';"
	response.write"opener.document.frm.ran_cnt.value = '"&ran_cnt&"';"
	response.write"opener.document.frm.work_man_cnt.value = '"&work_man_cnt&"';"
	response.write"opener.document.frm.alba_cnt.value = '"&alba_cnt&"';"
	'response.write"opener.document.getElementById('worK_ce').style.display = 'none';"
	'response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
