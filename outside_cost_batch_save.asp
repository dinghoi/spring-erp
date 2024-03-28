<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'	on Error resume next

	dim code_tab(100)
	slip_gubun = request.form("slip_gubun")
	sel_check = request.form("sel_check")+","

	i=1
	j= 1
	jj=0
	k=0
	do until i=0
		i=0
		i=instr(j,sel_check,",")'
	
		if	i=0 then
			exit do
		end if
		jj=i-1
		if j=i then
			code_tab(k)=""
	  	  else	  
			code_tab(k)=trim(mid(sel_check,j,jj-j+1))
		end if
		j=i+1
		k=k+1
	loop

	dbconn.BeginTrans

	j = 0
	for i = 0 to 100
		if code_tab(i) = "" then
			exit for
		end if
		j = j + 1
		slip_date = mid(code_tab(i),1,10)
		next_date = dateadd("m",1,slip_date)
		old_seq = mid(code_tab(i),11)

		Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&old_seq&"'"
		Set rs=DbConn.Execute(Sql)
	
		slip_gubun = rs("slip_gubun")
		customer = rs("customer")
		customer_no = rs("customer_no")
		emp_company = rs("emp_company")
		bonbu = rs("bonbu")
		saupbu = rs("saupbu")
		team = rs("team")
		org_name = rs("org_name")
		company = rs("company")
		account = rs("account")
		price = rs("price")
		cost = rs("cost")
		cost_vat = rs("cost_vat")
		if cost_vat <> 0 then
			vat_yn = "Y"
		  else
		  	vat_yn = "N"
		end if
		slip_memo = rs("slip_memo")
		emp_no = rs("emp_no")
		emp_name = rs("emp_name")
		emp_grade = rs("emp_grade")
		reg_id = rs("reg_id")
		rs.close()

		sql="select max(slip_seq) as max_seq from general_cost where slip_date='"&next_date&"'"
		set rs=dbconn.execute(sql)
			
		if	isnull(rs("max_seq"))  then
			slip_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			slip_seq = right(max_seq,3)
		end if
	
		sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,emp_company,bonbu,saupbu,team,org_name,company,account,account_item"&",pay_method,price,cost,vat_yn,cost_vat,customer,customer_no,emp_name,emp_no,emp_grade,pay_yn,slip_memo,tax_bill_yn,cancel_yn,end_yn,reg_id,reg_user,reg_date) values "&"('"&next_date&"','"&slip_seq&"','"&slip_gubun&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&company&"','"&account&"','"&account&"','현금',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&emp_name&"','"&emp_no&"','"&emp_grade&"','N','"&slip_memo&"','Y','N','N','"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)

		sql = "update general_cost set forward_yn='Y',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where slip_date='"&slip_date&"' and slip_seq = '"&old_seq&"'"
		dbconn.execute(sql)	  

	next

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

