<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

org_company=Request("org_company")
saupbu=Request("saupbu")
end_month=Request("end_month")
end_yn=Request("end_yn")
response.write(org_company)
	
cost_year = mid(end_month,1,4)
cost_month = mid(end_month,5)

from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
start_date = dateadd("m",-1,from_date)

response.write"<script language=javascript>"
response.write"alert('����ó����!!!');"
response.write"</script>"

dbconn.BeginTrans

sql = "select * from oil_unit where oil_unit_month = '"&end_month&"'"
Set rs_oil=DbConn.Execute(Sql)
if rs_oil.eof or rs_oil.bof then
	response.write"<script language=javascript>"
	response.write"alert('������ �ܰ��� �ԷµǾ� ���� �ʾ� ������ �� �� �����ϴ�.');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End
  else
' ������ �ܰ� �� ������ ���
	sql = "select * from transit_cost where (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (car_owner ='����') and (far > 0) and saupbu = '"&saupbu&"'"
	rs.Open sql, Dbconn, 1
	do until rs.eof

		if rs("team") = "������" or rs("team") = "SM1��" or rs("team") = "Repair��" or rs("team") = "SM2��" then
			oil_unit_id = "1"
		  else
			oil_unit_id = "2"
		end if
		sql = "select * from emp_master_month where emp_month = '"&end_month&"' and emp_no = '"&rs("mg_ce_id")&"'"
		Set rs_emp=DbConn.Execute(Sql)
'		response.write(sql)
'		response.write(rs_emp("emp_reside_company"))
		if rs_emp("emp_reside_company") = "��ȭȭ��" then
			liter = 8
		  else
			liter = 10
		end if
		rs_emp.close()

		if rs("oil_kind") = "����" then
			liter = 7
		end if
		
		sql = "select * from oil_unit where oil_unit_month = '"&end_month&"' and oil_unit_id = '"&oil_unit_id&"' and oil_kind = '"&rs("oil_kind")&"'"
		Set rs_etc=DbConn.Execute(Sql)
		oil_unit_average = rs_etc("oil_unit_average")
		rs_etc.close()
						
		oil_price = round(int(rs("far")) * oil_unit_average / liter)
		sql = "Update transit_cost set oil_unit="&oil_unit_average&", oil_price="&oil_price&" where mg_ce_id = '"&rs("mg_ce_id")&"' and run_date = '"&rs("run_date")&"' and run_seq ='"&rs("run_seq")&"'"
		dbconn.execute(sql)

		rs.movenext()
	loop
	rs.close()
' ������ �ܰ� �� ������ ��� ��

' ���κ� ��� ���� 
'	sql = "delete from person_cost where cost_month = '"&end_month&"' and saupbu = '"&saupbu&"'"
'	dbconn.execute(sql)

' ���� ������Ʈ
	sql = "select * from emp_master_month where emp_month = '"&end_month&"' and  emp_saupbu = '"&saupbu&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&start_date&"')"
	rs_emp.Open sql, Dbconn, 1
	emp_cnt = 0
	do until rs_emp.eof
		emp_cnt = emp_cnt + 1
		' �Ϲݺ�� 
'		sql = "update general_cost set emp_company='"&rs_emp("emp_company")&"',bonbu='"&rs_emp("emp_bonbu")&"',saupbu='"&rs_emp("emp_saupbu")&"',team='"&rs_emp("emp_team")&"',org_name='"&rs_emp("emp_org_name")&"',reside_place='"&rs_emp("emp_reside_place")&"' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (slip_gubun = '���') and (tax_bill_yn = 'N' or isnull(tax_bill_yn)) and (emp_no='"&rs_emp("emp_no")&"')"
'		dbconn.execute(sql)	  

		' �����
		sql = "update transit_cost set emp_company='"&rs_emp("emp_company")&"',bonbu='"&rs_emp("emp_bonbu")&"',saupbu='"&rs_emp("emp_saupbu")&"',team='"&rs_emp("emp_team")&"',org_name='"&rs_emp("emp_org_name")&"',reside_place='"&rs_emp("emp_reside_place")&"' where (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (mg_ce_id='"&rs_emp("emp_no")&"')"
		dbconn.execute(sql)	  

		' ��Ư��
		sql = "update overtime set emp_company='"&rs_emp("emp_company")&"',bonbu='"&rs_emp("emp_bonbu")&"',saupbu='"&rs_emp("emp_saupbu")&"',team='"&rs_emp("emp_team")&"',org_name='"&rs_emp("emp_org_name")&"',reside_place='"&rs_emp("emp_reside_place")&"' where (work_date >='"&from_date&"' and work_date <='"&to_date&"') and (mg_ce_id='"&rs_emp("emp_no")&"')"
		dbconn.execute(sql)	  

		' ī����ǥ
		sql = "update card_slip set emp_company='"&rs_emp("emp_company")&"',bonbu='"&rs_emp("emp_bonbu")&"',saupbu='"&rs_emp("emp_saupbu")&"',team='"&rs_emp("emp_team")&"',org_name='"&rs_emp("emp_org_name")&"',reside_place='"&rs_emp("emp_reside_place")&"',reside_company='"&rs_emp("emp_reside_company")&"' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no='"&rs_emp("emp_no")&"')"
		dbconn.execute(sql)	  
' ���� ������Ʈ ��

' ��� ���� üũ		
		if rs_emp("emp_end_date") = "1900-01-01" or isnull(rs_emp("emp_end_date")) or rs_emp("emp_end_date") >= from_date then
		  	emp_end = "�ٹ�"
		  else
			emp_end = "���"
		end if
		' �Ϲݺ��
		general_cnt = 0
		general_cost = 0
		general_pre_cnt = 0
		general_pre_cost = 0
		sql = "select pay_yn,count(slip_seq) as c_cnt,sum(cost) as cost from general_cost where (emp_no='"&rs_emp("emp_no")&"') "& _
		"and (slip_gubun = '���') and (tax_bill_yn = 'N' or isnull(tax_bill_yn)) and (cancel_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by pay_yn"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			if rs("pay_yn") = "N" then
				general_cnt = general_cnt + cint(rs("c_cnt"))
				general_cost = general_cost + cdbl(rs("cost"))
			  else
				general_pre_cnt = general_pre_cnt + cint(rs("c_cnt"))
				general_pre_cost = general_pre_cost + cdbl(rs("cost"))
			end if
			rs.movenext()
		loop
		rs.close()
		
		' ��Ư��
		overtime_cnt = 0
		overtime_cost = 0
		sql = "select cancel_yn,count(work_date) as c_cnt,sum(overtime_amt) as cost from overtime where (mg_ce_id='"&rs_emp("emp_no")&"') "& _
		"and (work_date >='"&from_date&"' and work_date <='"&to_date&"') and (cancel_yn = 'N') group by cancel_yn"
		'	response.write(sql)
		rs.Open sql, Dbconn, 1
		do until rs.eof
			overtime_cnt = overtime_cnt + cint(rs("c_cnt"))
			overtime_cost = overtime_cost + cdbl(rs("cost"))
			rs.movenext()
		loop
		rs.close()
		
		' �����
		gas_km = 0
		gas_unit = 0
		gas_cost = 0
		diesel_km = 0
		diesel_unit = 0
		diesel_cost = 0
		gasol_km = 0
		gasol_unit = 0
		gasol_cost = 0
		somopum_cost = 0
		fare_cnt = 0
		fare_cost = 0
		oil_cash_cost = 0
		repair_cost = 0
		repair_pre_cost = 0
		parking_cost = 0
		toll_cost = 0
		sql = "select * from transit_cost where (mg_ce_id='"&rs_emp("emp_no")&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cancel_yn = 'N')"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			if rs("car_owner") = "���߱���" then
				fare_cnt = fare_cnt + 1
				fare_cost = fare_cost + rs("fare")	
			end if
			if rs("car_owner") = "����" then
				if rs("oil_kind") = "�ֹ���" then
					gasol_km = gasol_km + rs("far")
				end if
				if rs("oil_kind") = "����" then
					diesel_km = diesel_km + rs("far")
				end if
				if rs("oil_kind") = "����" then
					gas_km = gas_km + rs("far")
				end if
			end if
			
			if rs("car_owner") = "ȸ��" then
				oil_cash_cost = oil_cash_cost + rs("oil_price")
				repair_cost = repair_cost + rs("repair_cost")
			end if
		
			parking_cost = parking_cost + rs("parking")
			toll_cost = toll_cost + rs("toll")
			rs.movenext()
		loop
		rs.close()
		if rs_emp("emp_team") = "������" or rs_emp("emp_team") = "Repair��" or rs_emp("emp_team") = "SM1��" or rs_emp("emp_team") = "SM2��" then
			oil_unit_id = "1"
		  else
			oil_unit_id = "2"
		end if
		
		sql = "select * from oil_unit where oil_unit_month = '"&end_month&"' and oil_unit_id = '"&oil_unit_id&"'"
'		response.write(sql)
		rs_etc.Open sql, Dbconn, 1
		do until rs_etc.eof
			if rs_etc("oil_kind") = "�ֹ���" then
				gasol_unit = rs_etc("oil_unit_average")
			  elseif rs_etc("oil_kind") = "����" then
				gas_unit = rs_etc("oil_unit_average")
			  else
				diesel_unit = rs_etc("oil_unit_average")
			end if	 
			rs_etc.movenext()
		loop
		rs_etc.close()
				
		if rs_emp("emp_reside_company") = "��ȭȭ��" then
			liter = 8
		  else
			liter = 10
		end if
		
		tot_km = gas_km + diesel_km + gasol_km
		somopum_cost = (tot_km) * 25
		
		gas_cost = round(gas_km * gas_unit / 7)
		diesel_cost = round(diesel_km * diesel_unit / liter)
		gasol_cost = round(gasol_km * gasol_unit / liter)
		tot_cost = gas_cost + diesel_cost + gasol_cost
		
		' ���� ī����
		juyoo_card_cnt = 0
		juyoo_card_cost = 0
		juyoo_card_cost_vat = 0
		juyoo_card_price = 0
		sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&rs_emp("emp_no")&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and card_type like '%����%'"
		
		Set rs = Dbconn.Execute (sql)
		if cint(rs("c_cnt")) <>  0 then
			juyoo_card_cnt = juyoo_card_cnt + cint(rs("c_cnt"))
			juyoo_card_cost = juyoo_card_cost + cdbl(rs("cost"))
			juyoo_card_cost_vat = juyoo_card_cost_vat + cdbl(rs("cost_vat"))
		end if
		rs.close()
		juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat

		' ī����
		card_cnt = 0
		card_cost = 0
		card_cost_vat = 0
		card_price = 0
		sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&rs_emp("emp_no")&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and card_type not like '%����%'"
'		sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&rs_emp("emp_no")&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
		
		Set rs = Dbconn.Execute (sql)
		if cint(rs("c_cnt")) <>  0 then
			card_cnt = card_cnt + cint(rs("c_cnt"))
			card_cost = card_cost + cdbl(rs("cost"))
			card_cost_vat = card_cost_vat + cdbl(rs("cost_vat"))
		end if
		rs.close()
		card_price = card_cost + card_cost_vat
		
		cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
		
' ���� ����
		sql = "select * from car_info where owner_emp_no ='"&rs_emp("emp_no")&"'"
		set rs_car=dbconn.execute(sql)
		if rs_car.eof then
			car_owner = "����"
		  else  	
			car_owner = rs_car("car_owner")
		end if	

'		if tot_km <> 0 then
		if car_owner = "����" then
			return_cash = cash_tot_cost - juyoo_card_price
		  else
			return_cash = cash_tot_cost
		end if
			
		sql = "select * from person_cost where cost_month ='"&end_month&"' and emp_no ='"&rs_emp("emp_no")&"'"
		set rs_person=dbconn.execute(sql)
		if rs_person.eof then
			variation_memo = ""
		  else  	
			variation_memo = rs_person("variation_memo")
		end if	
		rs_person.close()
		
		sql = "delete from person_cost where cost_month ='"&end_month&"' and emp_no ='"&rs_emp("emp_no")&"'"
		dbconn.execute(sql)
		
		sql = "insert into person_cost values ('"&end_month&"','"&rs_emp("emp_no")&"','"&rs_emp("emp_name")&"','"&rs_emp("emp_job")&"','"&emp_end&"','"&car_owner&"','"&rs_emp("emp_company")&"','"&rs_emp("emp_bonbu")&"','"&rs_emp("emp_saupbu")&"','"&rs_emp("emp_team")&"','"&rs_emp("emp_org_name")&"','"&rs_emp("emp_reside_place")&"','"&rs_emp("emp_reside_company")&"',"&general_cnt&","&general_cost&","&general_pre_cnt&","&general_pre_cost&","&overtime_cnt&","&overtime_cost&","&gas_km&","&gas_unit&","&gas_cost&","&diesel_km&","&diesel_unit&","&diesel_cost&","&gasol_km&","&gasol_unit&","&gasol_cost&","&tot_km&","&tot_cost&","&somopum_cost&","&fare_cnt&","&fare_cost&","&oil_cash_cost&","&repair_cost&","&repair_pre_cost&","&parking_cost&","&toll_cost&","&juyoo_card_cnt&","&juyoo_card_cost&","&juyoo_card_cost_vat&","&card_cnt&","&card_cost&","&card_cost_vat&","&return_cash&",'"&variation_memo&"',now())"
		dbconn.execute(sql)

'		if car_owner = "����" then
'			sql = "update card_slip set skip_yn='Y' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and emp_no ='"&rs_emp("emp_no")&"' "
'			dbconn.execute(sql)
'		end if
		rs_emp.movenext()
	loop
	rs_emp.close()
'���κ� ������� ��

' ���� �λ縶���� ���� ���� �ľ�
	if emp_cnt > 0 then
	
	' 4�뺸������ ��Ÿ �ΰǺ��� �˻�
		sql = "select * from insure_per where insure_year = '"&cost_year&"'"
		set rs_etc=dbconn.execute(sql)
		insure_tot_per = rs_etc("insure_tot_per")
		income_tax_per = rs_etc("income_tax_per")
		annual_pay_per = rs_etc("annual_pay_per")
		retire_pay_per = rs_etc("retire_pay_per")
		rs_etc.close()
		
		sql = "update org_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (saupbu = '"&saupbu&"')"
		dbconn.execute(sql)
	
	' �޿� SUM
		sql = "select pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_id,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where (pmg_saupbu = '"&saupbu&"') and (pmg_yymm ='"&end_month&"') and (pmg_id ='1') group by pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sort_seq = 0
			cost_detail = "�޿�"
	
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='"&cost_detail&"'"
			set rs_cost=dbconn.execute(sql)
			
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','�ΰǺ�','"&cost_detail&"',"&rs("tot_cost")&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("tot_cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='"&cost_detail&"'"
				dbconn.execute(sql)
			end if		
	' 2015-04-27
	' 4�뺸��� 
			insure_tot = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * insure_tot_per / 100)	
			sort_seq = 2
		
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='4�뺸��'"
			set rs_cost=dbconn.execute(sql)
			
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','�ΰǺ�','4�뺸��',"&insure_tot&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&insure_tot&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='4�뺸��'"
				dbconn.execute(sql)
			end if		
	
		' �ҵ漼 �������� 
			income_tax = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * income_tax_per / 100)		
			sort_seq = 3
		
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='�ҵ漼��������'"
			set rs_cost=dbconn.execute(sql)
			
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','�ΰǺ�','�ҵ漼��������',"&income_tax&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&income_tax&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='�ҵ漼��������'"
				dbconn.execute(sql)
			end if		
		' ��������
			annual_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * annual_pay_per / 100)		
			sort_seq = 4
		
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������'"
			set rs_cost=dbconn.execute(sql)
			
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','�ΰǺ�','��������',"&annual_pay&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&annual_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������'"
				dbconn.execute(sql)
			end if		
		' ��������
			retire_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * retire_pay_per / 100)		
			sort_seq = 5
		
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������'"
			set rs_cost=dbconn.execute(sql)
			
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','�ΰǺ�','��������',"&retire_pay&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&retire_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������'"
				dbconn.execute(sql)
			end if		
	
	' 2015-04-27 End
			rs.movenext()
		loop
		rs.close()
	' �� SUM
		sql = "select pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_id,sum(pmg_give_total) as cost from pay_month_give where (pmg_saupbu = '"&saupbu&"') and (pmg_yymm ='"&end_month&"') and (pmg_id ='2') group by pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_id"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sort_seq = 1
			cost_detail = "��"
	
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='"&cost_detail&"'"
			set rs_cost=dbconn.execute(sql)
			
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','�ΰǺ�','"&cost_detail&"',"&rs("cost")&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='"&cost_detail&"'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()
	
	' �˹ٺ�
		sql = "select company,bonbu,saupbu,team,org_name,sum(alba_give_total) as cost from pay_alba_cost where (saupbu = '"&saupbu&"') and (rever_yymm ='"&end_month&"') group by company,bonbu,saupbu,team,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='�˹ٺ�'"
			set rs_cost=dbconn.execute(sql)
		
			sort_seq = 8
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','�ΰǺ�','�˹ٺ�',"&rs("cost")&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�ΰǺ�' and cost_detail ='�˹ٺ�'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()
		
	'��Ư�� ����
		sql = "Update overtime set end_yn='Y' where work_date >= '"&from_date&"' and work_date <= '"&to_date&"' and saupbu ='"&saupbu&"'"
		dbconn.execute(sql)
	
	'�Ϲݺ��	
		sql = "Update general_cost set end_yn='Y' where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and saupbu ='"&saupbu&"'"
		dbconn.execute(sql)
	' DB SUM ó�� (���)
		sql = "select emp_company,bonbu,saupbu,team,org_name,account,sum(cost) as cost from general_cost where (slip_gubun = '���') and (cancel_yn = 'N') and (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof
						
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�Ϲݰ��' and cost_detail ='"&rs("account")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','�Ϲݰ��','"&rs("account")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�Ϲݰ��' and cost_detail ='"&rs("account")&"'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()
	' DB SUM ó�� (��� ��)
		sql = "select slip_gubun,emp_company,bonbu,saupbu,team,org_name,account,sum(cost) as cost from general_cost where (slip_gubun <> '���') and (cancel_yn = 'N') and (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by slip_gubun,emp_company,bonbu,saupbu,team,org_name,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			cost_id = rs("slip_gubun")
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&cost_id&"','"&rs("account")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()
	
	'�����
		sql = "Update transit_cost set end_yn='Y' where (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and saupbu ='"&saupbu&"'"
		dbconn.execute(sql)
	
	' DB SUM �����
		sql = "select emp_company,bonbu,saupbu,team,org_name,car_owner,sum(somopum+oil_price+fare+parking+toll) as cost from transit_cost where (cancel_yn = 'N') and (saupbu = '"&saupbu&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,car_owner"
		rs.Open sql, Dbconn, 1
		do until rs.eof
								
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�����' and cost_detail ='"&rs("car_owner")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','�����','"&rs("car_owner")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�����' and cost_detail ='"&rs("car_owner")&"'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()
	
	' DB SUM ����� (����������)
		sql = "select emp_company,bonbu,saupbu,team,org_name,sum(repair_cost) as cost from transit_cost where (cancel_yn = 'N') and (repair_cost > 0) and (saupbu = '"&saupbu&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
								
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�����' and cost_detail ='����������'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','�����','����������',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='�����' and cost_detail ='����������'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()

' ȸ�� ���� ���� ����ī�� ����
		sql = "select mg_ce_id from transit_cost where (car_owner = 'ȸ��') and (saupbu = '"&saupbu&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by mg_ce_id"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			sql = "update card_slip set com_drv_yn='Y' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no='"&rs("mg_ce_id")&"')"
			dbconn.execute(sql)	  

			rs.movenext()
		loop
		rs.close
	
	' ī���� ����
	'	sql = "select emp_company,bonbu,saupbu,team,org_name,reside_place,reside_company,account,sum(cost) as cost from card_slip where (end_sw = 'Y') and (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,reside_place,reside_company,account"
		sql = "select owner_company as emp_company,bonbu,saupbu,team,org_name,account,sum(cost) as cost from card_slip where (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (card_type not like '%����%' or com_drv_yn = 'Y')  group by owner_company,bonbu,saupbu,team,org_name,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof
								
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='����ī��' and cost_detail ='"&rs("account")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','����ī��','"&rs("account")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
	'			sum_cost = clng(rs("cost")) + clng(rs_cost(9+cost_month))
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='����ī��' and cost_detail ='"&rs("account")&"'"
				dbconn.execute(sql)
			end if		
			rs.movenext()
		loop
		rs.close()
	
	end if

		if end_yn = "C" then
			sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
			"' and saupbu = '"&saupbu&"'"
		  else
			sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
			"','"&saupbu&"','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
		end if
		dbconn.execute(sql)

	if emp_cnt = 0 then
		emp_msg = "�λ縶���� ������ ���� �ʾҽ��ϴ� "
	  else
		emp_msg = ""
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = emp_msg + "ó���� Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = emp_msg + "����ó�� �Ǿ����ϴ�...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
end if
rs_oil.close()
%>


