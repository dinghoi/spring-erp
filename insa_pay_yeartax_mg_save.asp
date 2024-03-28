<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	y_emp_no = request.form("emp_no")
	y_emp_name = request.form("emp_name")
	y_company = request.form("company_name")
	y_company_no = request.form("trade_no")
	emp_national = request.form("emp_national")
	
	'response.write(y_emp_no)
	'response.End
	
	y_year = request.form("inc_yyyy")
	
	y_householder = request.form("y_householder")
	Y_foreign = request.form("Y_foreign")
	y_disab = request.form("y_disab")
	y_woman = request.form("y_woman")
	y_single = request.form("y_single")
	y_blue = request.form("y_blue")
	y_live = request.form("y_live")
	y_change = request.form("y_change")
	
	y_total_pay =int(request.form("sum_give_tot"))
	y_total_bonus =int(request.form("sum_bunus_tot"))
	y_other_pay =int(request.form("sum_other_tot"))
	y_tax_no =int(request.form("sum_tax_no"))
	y_income_tax =int(request.form("sum_income_tax"))
	y_wetax =int(request.form("sum_wetax"))
	y_nps_amt =int(request.form("sum_nps_amt"))
	y_nhis_amt =int(request.form("sum_nhis_amt"))
	y_longcare_amt =int(request.form("sum_longcare_amt"))
	y_epi_amt =int(request.form("sum_epi_amt"))
	
	'response.write(y_woman)
	'response.End
	
	f_date = y_year + "-01" + "-01"
	y_to_date = y_year + "-12" + "-31"
'	response.write(wife_check)
'	response.end
	
	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_etc = Server.CreateObject("ADODB.Recordset")
    Set Rs_org = Server.CreateObject("ADODB.Recordset")
    Set Rs_emp = Server.CreateObject("ADODB.Recordset")
    Set Rs_fam = Server.CreateObject("ADODB.Recordset")
    Set Rs_year = Server.CreateObject("ADODB.Recordset")
    Set Rs_give = Server.CreateObject("ADODB.Recordset")
    Set Rs_dct = Server.CreateObject("ADODB.Recordset")
    Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
    Set Rs_ins = Server.CreateObject("ADODB.Recordset")
    Set Rs_sod = Server.CreateObject("ADODB.Recordset")
    Set RsCount = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans
	
' 입력화면에서 저장들어오면 연말정산 마스터..부양가족 삭제후 저장하는 방법으로...	
	
	
Sql = "select * from emp_master where emp_no = '"&y_emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
if not Rs_emp.eof then
   emp_in_date = rs_emp("emp_in_date")
   if emp_in_date < f_date then
          y_from_date = f_date
	  else
	      y_from_date = emp_in_date
   end if
   emp_birthday = rs_emp("emp_birthday")
   y_person_no1 = rs_emp("emp_person1")
   y_person_no2 = rs_emp("emp_person2")
   emp_nation_code = rs_emp("emp_nation_code")
   rs_emp.close()	
   emp_person = cstr(y_person_no1) + cstr(y_person_no2)	
   f_pseq = "01"
   if emp_nation_code = "001" then
          f_national = "1"
	  else
	      f_national = "9"
   end if

'부양가족 연말정산 마스터 등록 - 본인것도 부양가족db에 등록
           sql = "insert into pay_yeartax_family (f_year,f_emp_no,f_pseq,f_person_no,f_emp_name,f_family_name,f_rel,f_national,f_birthday,f_name,f_woman,f_single,f_disab) values "
		   sql = sql +	" ('"&y_year&"','"&y_emp_no&"','"&f_pseq&"','"&emp_person&"','"&y_emp_name&"','"&y_emp_name&"','본인','"&f_national&"','"&emp_birthday&"','"&y_emp_name&"','"&y_woman&"','"&y_single&"','"&y_disab&"')"
		
		   dbconn.execute(sql)	
	
   '부양가족
    y_age6_cnt = 0
	y_wife = 0
	y_old_cnt = 0
	y_age60_cnt = 0
	y_age20_cnt = 0
	y_pensioner_cnt = 0
	y_witak_cnt = 0
	y_disab_cnt = 0
	y_daja_cnt = 0
	y_holt_cnt = 0
	y_support_cnt = 0
    sql = "select * from emp_family where family_empno = '"&emp_no&"' ORDER BY family_empno,family_seq ASC"
    Rs_fam.Open Sql, Dbconn, 1
	do until Rs_fam.eof
		family_birthday = Rs_fam("family_birthday")
	    family_support_yn = Rs_fam("family_support_yn")
		family_live = Rs_fam("family_live")
		family_national = Rs_fam("family_national")
		if family_national = "내국인" then
               f_national = "1"
	       else
	           f_national = "9"
        end if
'		if family_live = "동거" then
'		       y_live = "거주자"
'		   else
'		       y_live = "비거주자"
'	    end if
		
		y_wife_chk = ""
	    y_old_chk = ""
	    y_age60_chk = ""
	    y_age20_chk = ""
		if family_support_yn = "Y" then '부양가족인 경우만
		   y_support_cnt = y_support_cnt + 1
		   if family_birthday < "1944-12-31" then     ' 추가공제 경로우대 70세이상
	              y_old_cnt = y_old_cnt + 1
				  y_old_chk = "Y"
		   end if 
		   family_children = Rs_fam("family_children")
		   if family_birthday > "2009-12-31" then     ' 자녀양육 6세이하
		          y_age6_cnt = y_age6_cnt + 1
		   end if 
		   family_rel = Rs_fam("family_rel")
		   if family_rel = "남편" or family_rel = "아내" then '기본공제 배우자공제
		          y_wife = 1
				  y_wife_chk = "Y"
		      else
		          y_wife = 0
		   end if  
		   if family_rel = "아들" or family_rel = "딸"  then  ' 추가공제 다자녀
	              y_daja_cnt = y_daja_cnt + 1
		   end if 
		   if family_birthday < "1954-12-31" then     ' 기본공제 60세이상
	              y_age60_cnt = y_age60_cnt + 1
				  y_age60_chk = "Y"
		   end if 
		   if family_birthday > "1994-01-01" then     ' 기본공제 20세이하
	              y_age20_cnt = y_age20_cnt + 1
				  y_age20_chk = "Y"
		   end if
		   family_pensioner = Rs_fam("family_pensioner")  
		   if family_pensioner = "Y" then ' 수급자
	              y_pensioner_cnt = y_pensioner_cnt + 1
		   end if
		   family_witak = Rs_fam("family_witak")
		   if family_witak = "Y" then                   ' 기본공제 위탁아동
	              y_witak_cnt = y_witak_cnt + 1
		   end if
	       family_disab = Rs_fam("family_disab")
	       family_merit = Rs_fam("family_merit")
	       family_serius = Rs_fam("family_serius")
		   if family_disab = "Y" or family_merit = "Y" or family_serius = "Y"  then  ' 추가공제 장애인
	              y_disab_cnt = y_disab_cnt + 1
		   end if
	       family_holt = Rs_fam("family_holt")  
		   if family_holt = "Y" then                   ' 기본공제 입양
	              y_holt_cnt = y_holt_cnt + 1
		   end if
		   family_name = Rs_fam("family_name")
		   family_person1 = Rs_fam("family_person1")
		   family_person2 = Rs_fam("family_person2")
		   family_person = cstr(family_person1) + cstr(family_person2)
		   
'부양가족 연말정산 마스터 등록 

           sql="select max(f_pseq) as max_seq from pay_yeartax_family where f_year='" & y_year & "' and f_emp_no='" & y_emp_no & "'"
		   set rs=dbconn.execute(sql)
		
		   if	isnull(rs("max_seq"))  then
			        f_pseq = "01"
		        else
			        max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			        f_pseq = right(max_seq,2)
		   end if

           sql = "insert into pay_yeartax_family (f_year,f_emp_no,f_pseq,f_person_no,f_emp_name,f_family_name,f_rel,f_national,f_birthday,f_name,f_wife,f_age20,f_age60,f_old,f_disab,f_merit,f_serius,f_pensioner,f_witak,f_holt,f_children) values "
		   sql = sql +	" ('"&y_year&"','"&y_emp_no&"','"&f_pseq&"','"&family_person&"','"&y_emp_name&"','"&family_name&"','"&family_rel&"','"&f_national&"','"&family_birthday&"','"&family_name&"','"&y_wife_chk&"','"&y_age20_chk&"','"&y_age60_chk&"','"&y_old_chk&"','"&family_disab&"','"&family_merit&"','"&family_serius&"','"&family_pensioner&"','"&family_witak&"','"&family_holt&"','"&family_children&"')"
		
		   dbconn.execute(sql)
	  end if
		 Rs_fam.movenext()
	loop
	Rs_fam.close()	

emp_user = request.cookies("nkpmg_user")("coo_user_name")

' 연말정산 소득자정보마스터 등록
		sql = "insert into pay_yeartax (y_year,y_emp_no,y_emp_name,y_person_no1,y_person_no2,y_company,y_company_no,y_from_date,y_to_date,y_national,y_live,y_change,y_householder,Y_foreign,y_disab,y_woman,y_single,y_blue,y_support_cnt,y_wife,y_age20_cnt,y_age60_cnt,y_pensioner_cnt,y_daja_cnt,y_holt_cnt,y_age6_cnt,y_old_cnt,y_disab_cnt,y_total_pay,y_total_bonus,y_other_pay,y_tax_no,y_income_tax,y_wetax,y_nps_amt,y_nhis_amt,y_epi_amt,y_longcare_amt) values "
		sql = sql +	" ('"&y_year&"','"&y_emp_no&"','"&y_emp_name&"','"&y_person_no1&"','"&y_person_no2&"','"&y_company&"','"&y_company_no&"','"&y_from_date&"','"&y_to_date&"','"&emp_national&"','"&y_live&"','"&y_change&"','"&y_householder&"','"&Y_foreign&"','"&y_disab&"','"&y_woman&"','"&y_single&"','"&y_blue&"','"&y_support_cnt&"','"&y_wife&"','"&y_age20_cnt&"','"&y_age60_cnt&"','"&y_pensioner_cnt&"','"&y_daja_cnt&"','"&y_holt_cnt&"','"&y_age6_cnt&"','"&y_old_cnt&"','"&y_disab_cnt&"','"&y_total_pay&"','"&y_total_bonus&"','"&y_other_pay&"','"&y_tax_no&"','"&y_income_tax&"','"&y_wetax&"','"&y_nps_amt&"','"&y_nhis_amt&"','"&y_epi_amt&"','"&y_longcare_amt&"')"

		dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	  else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if
  
  else
    end_msg = "등록된 직원이 아닙니다...."
end if
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"self.opener.location.reload();"	
	response.write"location.replace('insa_pay_yeartax_mg.asp');"	
	'response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
