<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	dim cc_tab(20,20)
	dim dd_tab(20,11)
	dim mm_tab(20,10)
	dim ii_tab(20,10)
	
	emp_no = request.form("emp_no1")
	inc_yyyy = request.form("inc_yyyy")
	
'	response.write(inc_yyyy)
'	response.write(emp_no)

for i = 1 to 20
    cc_tab(i,1) = ""
	cc_tab(i,2) = ""
	cc_tab(i,3) = ""
	cc_tab(i,4) = ""
	cc_tab(i,5) = ""
	
	cc_tab(i,6) = 0
	cc_tab(i,7) = 0
	cc_tab(i,8) = 0
	cc_tab(i,9) = 0
	cc_tab(i,10) = 0
	cc_tab(i,11) = 0
	
	cc_tab(i,12) = 0
	cc_tab(i,13) = 0
	cc_tab(i,14) = 0
	cc_tab(i,15) = 0
	cc_tab(i,16) = 0
	cc_tab(i,17) = 0
	
	cc_tab(i,18) = 0
	cc_tab(i,19) = 0
	cc_tab(i,20) = 0
next

for i = 1 to 20
    dd_tab(i,1) = ""
	dd_tab(i,2) = ""
	dd_tab(i,3) = ""
	
	dd_tab(i,4) = 0
	dd_tab(i,5) = 0
	dd_tab(i,6) = 0
	dd_tab(i,7) = 0
	dd_tab(i,8) = 0
	dd_tab(i,9) = 0
	dd_tab(i,10) = 0
	dd_tab(i,11) = 0
next	

for i = 1 to 20
    mm_tab(i,1) = ""
	mm_tab(i,2) = ""
	mm_tab(i,3) = ""
	
	mm_tab(i,4) = 0
	mm_tab(i,5) = 0
	mm_tab(i,6) = 0
	mm_tab(i,7) = 0
	mm_tab(i,8) = 0
	mm_tab(i,9) = 0
	mm_tab(i,10) = 0
next	

for i = 1 to 20
    ii_tab(i,1) = ""
	ii_tab(i,2) = ""
	ii_tab(i,3) = ""
	
	ii_tab(i,4) = 0
	ii_tab(i,5) = 0
	ii_tab(i,6) = 0
	ii_tab(i,7) = 0
	ii_tab(i,8) = 0
	ii_tab(i,9) = 0
	ii_tab(i,10) = 0
next	
	
Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")

Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_ann = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_dona = Server.CreateObject("ADODB.Recordset")
Set rs_duct = Server.CreateObject("ADODB.Recordset")
Set rs_cred = Server.CreateObject("ADODB.Recordset")
Set rs_hous = Server.CreateObject("ADODB.Recordset")
Set rs_houm = Server.CreateObject("ADODB.Recordset")
Set rs_savi = Server.CreateObject("ADODB.Recordset")
Set rs_other = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_first_date = rs_emp("emp_first_date")
		emp_in_date = rs_emp("emp_in_date")
		emp_end_date = rs_emp("emp_end_date")
		emp_type = rs_emp("emp_type")
		emp_grade = rs_emp("emp_grade")
		emp_position = rs_emp("emp_position")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
		emp_reside_place = rs_emp("emp_reside_place")
		emp_reside_company = rs_emp("emp_reside_company")
		emp_disabled = rs_emp("emp_disabled")
		emp_disab_grade = rs_emp("emp_disab_grade")
   else
		emp_first_date = ""
		emp_in_date = ""
		emp_end_date = ""
		emp_type = ""
		emp_grade = ""
		emp_position = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
		emp_reside_place = ""
		emp_reside_company = ""
		emp_disabled = ""
		emp_disab_grade = ""
end if

t_year = inc_yyyy





'신용카드
sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
if rs_year("y_final") = "N" then
	
sql = " SELECT c_year,c_emp_no,c_person_no,c_rel,cc_name,count(*) as cc_count" & _
			"   from pay_yeartax_credit " & _
            "   WHERE c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' " & _
			"   group by c_year,c_emp_no,c_person_no,c_rel,cc_name " & _
			"   order by c_emp_no,c_person_no,c_id,c_seq ASC "
rs_cred.Open Sql, Dbconn, 1
i = 0
do until rs_cred.eof
       i = i + 1
	          cc_tab(i,1) = rs_cred("c_year")
	          cc_tab(i,2) = rs_cred("c_emp_no")
	          cc_tab(i,3) = rs_cred("c_person_no")
	          cc_tab(i,4) = rs_cred("c_rel")
	          cc_tab(i,5) = rs_cred("cc_name")
	rs_cred.MoveNext()
loop
rs_cred.close()	

sql = "select * from pay_yeartax_credit where c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' ORDER BY c_emp_no,c_person_no,c_id,c_seq ASC"
rs_cred.Open Sql, Dbconn, 1
do until rs_cred.eof
   for i = 1 to 20
	   if rs_cred("c_year") = cc_tab(i,1) and rs_cred("c_emp_no") = cc_tab(i,2) and rs_cred("c_person_no") = cc_tab(i,3) then
		   if rs_cred("c_id") = "신용카드" and rs_cred("c_market")  = "Y" then 
		         cc_tab(i,8) =  cc_tab(i,8) + rs_cred("c_nts_amt")   
				 cc_tab(i,9) =  cc_tab(i,9) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "신용카드" and rs_cred("c_transit")  = "Y" then 
		         cc_tab(i,10) =  cc_tab(i,10) + rs_cred("c_nts_amt")   
				 cc_tab(i,11) =  cc_tab(i,11) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "신용카드" and rs_cred("c_transit")  <> "Y" and rs_cred("c_transit") <> "Y" then 
		         cc_tab(i,6) =  cc_tab(i,6) + rs_cred("c_nts_amt")   
				 cc_tab(i,7) =  cc_tab(i,7) + rs_cred("c_other_amt")   
           end if
		   
		   if rs_cred("c_id") = "직불카드" and rs_cred("c_market")  = "Y" then 
		         cc_tab(i,14) =  cc_tab(i,14) + rs_cred("c_nts_amt")   
				 cc_tab(i,15) =  cc_tab(i,15) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "직불카드" and rs_cred("c_transit")  = "Y" then 
		         cc_tab(i,16) =  cc_tab(i,16) + rs_cred("c_nts_amt")   
				 cc_tab(i,17) =  cc_tab(i,17) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "직불카드" and rs_cred("c_transit")  <> "Y" and rs_cred("c_transit") <> "Y" then 
		         cc_tab(i,12) =  cc_tab(i,12) + rs_cred("c_nts_amt")   
				 cc_tab(i,13) =  cc_tab(i,13) + rs_cred("c_other_amt")   
           end if
		   
		   if rs_cred("c_id") = "현금영수증" and rs_cred("c_market")  = "Y" then 
		         cc_tab(i,19) =  cc_tab(i,14) + rs_cred("c_nts_amt")   
           end if
		   if rs_cred("c_id") = "현금영수증" and rs_cred("c_transit")  = "Y" then 
		         cc_tab(i,20) =  cc_tab(i,16) + rs_cred("c_nts_amt")   
           end if
		   if rs_cred("c_id") = "현금영수증" and rs_cred("c_transit")  <> "Y" and rs_cred("c_transit") <> "Y" then 
		         cc_tab(i,18) =  cc_tab(i,12) + rs_cred("c_nts_amt")   
           end if
       end if
	next
	rs_cred.MoveNext()
loop
rs_cred.close()	

'기부금
sql = " SELECT d_year,d_emp_no,d_person_no,count(*) as dd_count " & _
			"   from pay_yeartax_donation " & _
            "   WHERE d_year = '"&inc_yyyy&"' and d_emp_no = '"&emp_no&"' " & _
			"   group by d_year,d_emp_no,d_person_no " & _
			"   order by d_emp_no,d_person_no,d_seq ASC "
rs_dona.Open Sql, Dbconn, 1
i = 0
do until rs_dona.eof
       i = i + 1
	          dd_tab(i,1) = rs_dona("d_year")
	          dd_tab(i,2) = rs_dona("d_emp_no")
	          dd_tab(i,3) = rs_dona("d_person_no")
	rs_dona.MoveNext()
loop
rs_dona.close()		

sql = "select * from pay_yeartax_donation where d_year = '"&inc_yyyy&"' and d_emp_no = '"&emp_no&"' ORDER BY d_emp_no,d_person_no,d_seq ASC"
rs_dona.Open Sql, Dbconn, 1
do until rs_dona.eof
   for i = 1 to 20
	   if rs_dona("d_year") = dd_tab(i,1) and rs_dona("d_emp_no") = dd_tab(i,2) and rs_dona("d_person_no") = dd_tab(i,3) then
		   if rs_dona("d_data_gubun") = "정치자금기부금" and rs_dona("d_nts_chk")  = "Y" then 
		         if rs_dona("d_amt") > 100000 then
				      dd_tab(i,6) =  dd_tab(i,6) + rs_dona("d_amt")   
					else
					  dd_tab(i,4) =  dd_tab(i,4) + rs_dona("d_amt")   
				 end if
			end if
			if rs_dona("d_data_gubun") = "정치자금기부금" and rs_dona("d_nts_chk")  <> "Y" then 
			     if rs_dona("d_amt") > 100000 then
			          dd_tab(i,7) =  dd_tab(i,7) + rs_dona("d_amt")   
					else  
					  dd_tab(i,5) =  dd_tab(i,5) + rs_dona("d_amt")   
				 end if
           end if
		   if rs_dona("d_data_gubun") = "법정기부금" and rs_dona("d_nts_chk")  = "Y" then 
		         dd_tab(i,8) =  dd_tab(i,8) + rs_dona("d_amt")   
		   end if
		   if rs_dona("d_data_gubun") = "법정기부금" and rs_dona("d_nts_chk")  <> "Y" then 
			     dd_tab(i,9) =  dd_tab(i,9) + rs_dona("d_amt")   
           end if
		   if rs_dona("d_data_gubun") = "종교단체외지정기부금" and rs_dona("d_nts_chk")  = "Y" then 
		         dd_tab(i,10) =  dd_tab(i,10) + rs_dona("d_amt")   
		   end if
		   if rs_dona("d_data_gubun") = "종교단체외지정기부금" and rs_dona("d_nts_chk")  <> "Y" then 
			     dd_tab(i,11) =  dd_tab(i,11) + rs_dona("d_amt")   
           end if
		   if rs_dona("d_data_gubun") = "종교단체지정기부금" and rs_dona("d_nts_chk")  = "Y" then 
		         dd_tab(i,10) =  dd_tab(i,10) + rs_dona("d_amt")   
		   end if
		   if rs_dona("d_data_gubun") = "종교단체지정기부금" and rs_dona("d_nts_chk")  <> "Y" then 
			     dd_tab(i,11) =  dd_tab(i,11) + rs_dona("d_amt")   
           end if
		   if rs_dona("d_data_gubun") = "우리사주조합기부금" and rs_dona("d_nts_chk")  = "Y" then 
		         dd_tab(i,10) =  dd_tab(i,10) + rs_dona("d_amt")   
		   end if
		   if rs_dona("d_data_gubun") = "우리사주조합기부금" and rs_dona("d_nts_chk")  <> "Y" then 
			     dd_tab(i,11) =  dd_tab(i,11) + rs_dona("d_amt")   
           end if
       end if
	next
	rs_dona.MoveNext()
loop
rs_dona.close()	

'의료비
sql = " SELECT m_year,m_emp_no,m_person_no,count(*) as mm_count " & _
			"   from pay_yeartax_medical " & _
            "   WHERE m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' " & _
			"   group by m_year,m_emp_no,m_person_no " & _
			"   order by m_emp_no,m_person_no,m_seq ASC "
rs_medi.Open Sql, Dbconn, 1
i = 0
do until rs_medi.eof
       i = i + 1
	          mm_tab(i,1) = rs_medi("m_year")
	          mm_tab(i,2) = rs_medi("m_emp_no")
	          mm_tab(i,3) = rs_medi("m_person_no")
	rs_medi.MoveNext()
loop
rs_medi.close()		

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
rs_medi.Open Sql, Dbconn, 1
do until rs_medi.eof
   for i = 1 to 20
	   if rs_medi("m_year") = mm_tab(i,1) and rs_medi("m_emp_no") = mm_tab(i,2) and rs_medi("m_person_no") = mm_tab(i,3) then
		   if rs_medi("m_data_gubun") = "국세청" then 
		         mm_tab(i,4) =  mm_tab(i,4) + rs_medi("m_amt")   
		      else
			     mm_tab(i,5) =  mm_tab(i,5) + rs_medi("m_amt")   
           end if
       end if
	next
	rs_medi.MoveNext()
loop
rs_medi.close()	

'보험료
sql = " SELECT i_year,i_emp_no,i_person_no,count(*) as ii_count " & _
			"   from pay_yeartax_insurance " & _
            "   WHERE i_year = '"&inc_yyyy&"' and i_emp_no = '"&emp_no&"' " & _
			"   group by i_year,i_emp_no,i_person_no " & _
			"   order by i_emp_no,i_person_no,i_seq ASC "
rs_ins.Open Sql, Dbconn, 1
i = 0
do until rs_ins.eof
       i = i + 1
	          ii_tab(i,1) = rs_ins("i_year")
	          ii_tab(i,2) = rs_ins("i_emp_no")
	          ii_tab(i,3) = rs_ins("i_person_no")
	rs_ins.MoveNext()
loop
rs_ins.close()		

sql = "select * from pay_yeartax_insurance where i_year = '"&inc_yyyy&"' and i_emp_no = '"&emp_no&"' ORDER BY i_emp_no,i_person_no,i_seq ASC"
rs_ins.Open Sql, Dbconn, 1
do until rs_ins.eof
   for i = 1 to 20
	   if rs_ins("i_year") = ii_tab(i,1) and rs_ins("i_emp_no") = ii_tab(i,2) and rs_ins("i_person_no") = ii_tab(i,3) then
		   if rs_ins("i_disab_chk") = "Y" then 
		         ii_tab(i,4) =  ii_tab(i,4) + rs_ins("i_nts_amt")
				 ii_tab(i,5) =  ii_tab(i,5) + rs_ins("i_other_amt")     
		      else
			     ii_tab(i,6) =  ii_tab(i,6) + rs_ins("i_nts_amt")   
				 ii_tab(i,7) =  ii_tab(i,7) + rs_ins("i_other_amt")   
           end if
       end if
	next
	rs_ins.MoveNext()
loop
rs_ins.close()	

dbconn.BeginTrans

'교육비
sql = " SELECT e_year,e_emp_no,e_person_no,count(*) as ee_count, " & _
			"   sum(e_nts_amt) as e_nts_amt,sum(e_other_amt) as e_other_amt " & _
			"   from pay_yeartax_edu " & _
            "   WHERE e_year = '"&inc_yyyy&"' and e_emp_no = '"&emp_no&"' " & _
			"   group by e_year,e_emp_no,e_person_no " & _
			"   order by e_emp_no,e_person_no,e_seq ASC "
rs_edu.Open Sql, Dbconn, 1
ok_e = 0
do until rs_edu.eof
    e_year = rs_edu("e_year")
	e_emp_no = rs_edu("e_emp_no")
	e_person_no = rs_edu("e_person_no")
	ee_nts_amt = clng(rs_edu("e_nts_amt"))
	ee_other_amt = clng(rs_edu("e_other_amt"))
		   
	sql = "Update pay_yeartax_family set e_nts_amt='"&ee_nts_amt&"',e_other_amt='"&ee_other_amt&"' where f_year = '"&e_year&"' and f_emp_no = '"&e_emp_no&"' and f_person_no = '"&e_person_no&"'"
		   
	dbconn.execute(sql)
		   
	ok_e = ok_e + 1 
	
	rs_edu.MoveNext()
loop
rs_edu.close()	

ok_c = 0
for i = 1 to 20
    if cc_tab(i,2) = "" or isnull(cc_tab(i,2)) then 
		   exit for
	   else 
		   c_year = cc_tab(i,1)
		   c_emp_no = cc_tab(i,2)
		   c_person_no = cc_tab(i,3)
		   
		   nts_market = cc_tab(i,8) + cc_tab(i,14) + cc_tab(i,19) 
		   nts_transit = cc_tab(i,10) + cc_tab(i,16) + cc_tab(i,20)
		   other_market = cc_tab(i,9) + cc_tab(i,15)
		   other_transit = cc_tab(i,11) + cc_tab(i,17)
		   nts_hap = cc_tab(i,6) + cc_tab(i,12) + cc_tab(i,18) + nts_market + nts_transit
		   other_hap = cc_tab(i,7) + cc_tab(i,13) + other_market + other_transit
		   
		   sql = "Update pay_yeartax_family set c_credit_nts='"&cc_tab(i,6)&"',c_credit_other='"&cc_tab(i,7)&"',c_cash_nts='"&cc_tab(i,18)&"',c_direct_nts='"&cc_tab(i,12)&"',c_direct_other='"&cc_tab(i,13)&"',c_market_nts='"&nts_market&"',c_market_other='"&other_market&"',c_transit_nts='"&nts_transit&"',c_transit_other='"&other_transit&"' where f_year = '"&c_year&"' and f_emp_no = '"&c_emp_no&"' and f_person_no = '"&c_person_no&"'"
		   
	       dbconn.execute(sql)
		   
		   ok_c = ok_c + 1 
    end if
next		   

ok_d = 0
for i = 1 to 20
    if dd_tab(i,2) = "" or isnull(dd_tab(i,2)) then 
		   exit for
	   else 
		   d_year = dd_tab(i,1)
		   d_emp_no = dd_tab(i,2)
		   d_person_no = dd_tab(i,3)
		   
		   sql = "Update pay_yeartax_family set d_poli_nts='"&dd_tab(i,4)&"',d_poli_other='"&dd_tab(i,5)&"',d_poli10_nts='"&dd_tab(i,6)&"',d_poli10_other='"&dd_tab(i,7)&"',d_law_nts='"&dd_tab(i,8)&"',d_law_other='"&dd_tab(i,9)&"',d_ji_nts='"&dd_tab(i,10)&"',d_ji_other='"&dd_tab(i,11)&"' where f_year = '"&d_year&"' and f_emp_no = '"&d_emp_no&"' and f_person_no = '"&d_person_no&"'"
		   
	       dbconn.execute(sql)
		   
		   ok_d = ok_d + 1 
    end if
next		  

ok_m = 0
for i = 1 to 20
    if mm_tab(i,2) = "" or isnull(mm_tab(i,2)) then 
		   exit for
	   else 
		   m_year = mm_tab(i,1)
		   m_emp_no = mm_tab(i,2)
		   m_person_no = mm_tab(i,3)
		   
		   sql = "Update pay_yeartax_family set m_nts_amt='"&mm_tab(i,4)&"',m_other_amt='"&mm_tab(i,5)&"' where f_year = '"&m_year&"' and f_emp_no = '"&m_emp_no&"' and f_person_no = '"&m_person_no&"'"
		   
	       dbconn.execute(sql)
		   
		   ok_m = ok_m + 1 
    end if
next		  

ok_i = 0
for i = 1 to 20
    if ii_tab(i,2) = "" or isnull(ii_tab(i,2)) then 
		   exit for
	   else 
		   i_year = ii_tab(i,1)
		   i_emp_no = ii_tab(i,2)
		   i_person_no = ii_tab(i,3)
		   
		   sql = "Update pay_yeartax_family set i_ilban_nts='"&ii_tab(i,6)&"',i_ilban_other='"&ii_tab(i,7)&"',i_disab_nts='"&ii_tab(i,4)&"',i_disab_other='"&ii_tab(i,5)&"' where f_year = '"&i_year&"' and f_emp_no = '"&i_emp_no&"' and f_person_no = '"&i_person_no&"'"
		   
	       dbconn.execute(sql)
		   
		   ok_i = ok_i + 1 
    end if
next		  

    sql = "Update pay_yeartax set y_final='Y' where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
		   
	dbconn.execute(sql)

		   
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "제출중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "제출되었습니다...."
	end if


else
    end_msg = "이미 확정제출 하셨습니다...."
end if


	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"..."&ok_c&"..."&ok_d&"..."&ok_e&"..."&ok_m&"..."&ok_i&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
	