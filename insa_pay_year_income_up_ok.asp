<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<% 
'	on Error resume next

	incom_year = request.form("incom_year")
	incom_company = request.form("incom_company")
'	pmg_date = request.form("pmg_date")
	objFile = request.form("objFile")
	
	w_cnt = 0

    emp_user = request.cookies("nkpmg_user")("coo_user_name")

	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_org = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
	Set Rs_year = Server.CreateObject("ADODB.Recordset")
	Set Rs_ins = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect
	
'국민연금 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&incom_year&"' and insu_id = '5501' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nps_emp = formatnumber(rs_ins("emp_rate"),3)
		nps_com = formatnumber(rs_ins("com_rate"),3)
		nps_from = rs_ins("from_amt")
		nps_to = rs_ins("to_amt")
   else
		nps_emp = 0
		nps_com = 0
		nps_from = 0
		nps_to = 0
end if
rs_ins.close()

'건강보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&incom_year&"' and insu_id = '5502' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nhis_emp = formatnumber(rs_ins("emp_rate"),3)
		nhis_com = formatnumber(rs_ins("com_rate"),3)
		nhis_from = rs_ins("from_amt")
		nhis_to = rs_ins("to_amt")
   else
		nhis_emp = 0  
		nhis_com = 0
		nhis_from = 0
		his_to = 0
end if
rs_ins.close()		
	
	
	dbconn.BeginTrans

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"
	
	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
    if rowcount > -1 then
		for i=0 to rowcount
			if xgr(1,i) = "" or isnull(xgr(1,i)) then
				exit for
			end if
	        incom_year = xgr(0,i)
	' 사번체크 				
			Sql = "select * from emp_master where emp_no = '"&xgr(1,i)&"'"
			Set rs_emp = DbConn.Execute(SQL)
			if rs_emp.eof then
                emp_no = xgr(1,i)
				emp_name = xgr(2,i)
				emp_in_date = ""
				emp_grade = ""
				emp_type = "정직"
				emp_pay_type = "1"
				emp_company = ""
				emp_org_code = ""
				emp_org_name = ""
			  else
				emp_no = xgr(1,i)
				emp_name = rs_emp("emp_name")	
				emp_company = rs_emp("emp_company")	
				emp_bonbu = rs_emp("emp_bonbu")	
				emp_saupbu = rs_emp("emp_saupbu")	
				emp_team = rs_emp("emp_team")	
				emp_org_code = rs_emp("emp_org_code")	
				emp_org_name = rs_emp("emp_org_name")	
				emp_reside_company = rs_emp("emp_reside_company")	
				emp_in_date = rs_emp("emp_in_date")
				emp_grade = rs_emp("emp_grade")
				emp_position = rs_emp("emp_position")
				emp_type = rs_emp("emp_type")	  
				emp_pay_type = rs_emp("emp_pay_type")	  
				cost_center = rs_emp("cost_center")	  
				cost_group = rs_emp("cost_group")	  
			end if
			w_cnt = w_cnt + 1
            
			incom_total_pay = xgr(3,i)
	'		incom_base_pay = xgr(4,i)
	'		incom_overtime_pay = xgr(5,i)
	'		incom_meals_pay = xgr(6,i)
	'		incom_severance_pay = xgr(7,i)
	'		incom_month_amount = xgr(8,i)
	 ' 기본급등 계산		
			mon13_pay = int(incom_total_pay / 13)
			meals_pay = 100000
			ot_pay = int((mon13_pay - meals_pay) * 0.09)
			base_pay = int(mon13_pay - meals_pay - ot_pay)
			mon_amt = base_pay + ot_pay
			se_pay = int(mon13_pay)
					
			incom_base_pay = base_pay
			incom_overtime_pay = ot_pay
			incom_meals_pay = meals_pay
			incom_severance_pay = se_pay
			incom_month_amount = mon_amt
			 		
	 '국민연금 계산
	        incom_nps_amount = xgr(9,i)
	'		incom_nps = xgr(10,i)
            nps_amt = incom_nps_amount * (nps_emp / 100)
            nps_amt = int(nps_amt)
            incom_nps = (int(nps_amt / 10)) * 10

     '건강보험 계산
	        incom_nhis_amount = xgr(11,i)
	'		incom_nhis = xgr(12,i)
            nhis_amt = incom_nhis_amount * (nhis_emp / 100)
            nhis_amt = int(nhis_amt)
            incom_nhis = (int(nhis_amt / 10)) * 10
		   
	 '항목	   
		    incom_go_yn = xgr(13,i)
	        incom_san_yn = xgr(14,i)
	        incom_long_yn = xgr(15,i)
	        incom_incom_yn = xgr(16,i)
	        incom_family_cnt = xgr(17,i)
	        incom_wife_yn = xgr(18,i)
	        incom_age20 = xgr(19,i)
	        incom_age60 = xgr(20,i)
	        incom_old = xgr(21,i)
	        incom_disab = xgr(22,i)
	        incom_woman = xgr(23,i)
	        incom_retirement_bank = xgr(24,i)
			incom_first3_percent = 0
		   

		Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&incom_year&"'"
		set Rs_year=dbconn.execute(sql)				
		if Rs_year.eof or Rs_year.bof then
			
			sql = "insert into pay_year_income (incom_emp_no,incom_year,incom_emp_name,incom_in_date,incom_grade,incom_emp_type,incom_pay_type,incom_company,incom_org_code,incom_org_name,incom_base_pay,incom_overtime_pay,incom_meals_pay,incom_severance_pay,incom_total_pay,incom_first3_percent,incom_month_amount,incom_nps_amount,incom_nhis_amount,incom_family_cnt,incom_nps,incom_nhis,incom_go_yn,incom_san_yn,incom_long_yn,incom_incom_yn,incom_wife_yn,incom_age20,incom_age60,incom_old,incom_disab,incom_woman,incom_retirement_bank,incom_reg_date,incom_reg_user) values "
		sql = sql +	" ('"&emp_no&"','"&incom_year&"','"&emp_name&"','"&emp_in_date&"','"&emp_grade&"','"&emp_type&"','"&emp_pay_type&"','"&emp_company&"','"&emp_org_code&"','"&emp_org_name&"','"&incom_base_pay&"','"&incom_overtime_pay&"','"&incom_meals_pay&"','"&incom_severance_pay&"','"&incom_total_pay&"','"&incom_first3_percent&"','"&incom_month_amount&"','"&incom_nps_amount&"','"&incom_nhis_amount&"','"&incom_family_cnt&"','"&incom_nps&"','"&incom_nhis&"','"&incom_go_yn&"','"&incom_san_yn&"','"&incom_long_yn&"','"&incom_incom_yn&"','"&incom_wife_yn&"','"&incom_age20&"','"&incom_age60&"','"&incom_old&"','"&incom_disab&"','"&incom_woman&"','"&incom_retirement_bank&"',now(),'"&emp_user&"')"
		
		    dbconn.execute(sql)
		end if
		
	
		next
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = cstr(w_cnt) +" 건 등록 완료되었습니다...."
	end if

	'err_msg = cstr(rowcount+1) + " 건 처리되었습니다..."
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('insa_pay_year_income_mg.asp');"
	response.write"</script>"
	Response.End

	rs.close
	cn.close
	rs_etc.close
	set rs = nothing
	set cn = nothing
	set rs_etc = nothing
%>