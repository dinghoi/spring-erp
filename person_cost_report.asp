<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/query/person/query_person_cost.asp" -->
<%
on Error resume next

Dim from_date
Dim to_date
Dim win_sw

cost_month=Request.form("cost_month")

sql = "select * from emp_master_month where emp_month = '"&cost_month&"' and emp_no = '"&user_id&"'"
'Response.write sql
set rs=dbconn.execute(sql)
if rs.eof or rs.bof then
	month_check = "N"
else
	emp_company = rs("emp_company")
	bunbu = rs("emp_bonbu")
	saupbu = rs("emp_saupbu")
	team = rs("emp_team")
	org_name = rs("emp_org_name")
	reside_place = rs("emp_reside_place")
	reside_company = rs("emp_reside_company")
end if
rs.close()

if saupbu = ""  or isnull(saupbu) then
	saupbu = "����οܳ�����"
end if

if cost_month = "" then
'	cost_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	sql = "select max(end_month) as end_month from cost_end where saupbu = '"&saupbu&"'"
	set rs=dbconn.execute(sql)
	if rs("end_month") = "" or isnull(rs("end_month")) then
		cost_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	else
		imsi_date = mid(rs("end_month"),1,4) + "-" + mid(rs("end_month"),5,2) + "-01"
		end_date = datevalue(imsi_date)
		end_date = dateadd("m",1,end_date)
		cost_month = mid(end_date,1,4) + mid(end_date,6,2)
	end if
	rs.close()
end If

from_date = mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

be_yy = int(mid(cost_month,1,4))
be_mm = int(mid(cost_month,5) - 1)
if be_mm = 0 then
	be_month = cstr(be_yy - 1) + "12"
else
  	be_month = cstr(be_yy) + right("0" + cstr(be_mm),2)
end if

' ����� ���� ���� ����Ÿ ����
	sql = "select * from transit_cost where (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (emp_company ='' or isnull(emp_company)) and mg_ce_id ='"&user_id&"'"
	rs.Open sql, Dbconn, 1
	do until rs.eof

		sql = "Update transit_cost set emp_company='"&emp_company&"', bonbu='"&bonbu&"', saupbu='"&saupbu&"', team='"&team&"', org_name='"&org_name&"', reside_place='"&reside_place&"' where mg_ce_id = '"&rs("mg_ce_id")&"' and run_date = '"&rs("run_date")&"' and run_seq ='"&rs("run_seq")&"'"
		dbconn.execute(sql)
		rs.movenext()
	loop
	rs.close()
' ����� ���� ���� ����Ÿ ���� ��

' ���� ���� (������ 2���̻��̸�.. �հ� �̻��Ѱ�..)
sql = "select * from car_info where owner_emp_no ='"&user_id&"' "
set rs_car=dbconn.execute(sql)
if rs_car.eof then
	car_info = "��������"
	car_owner = "����"
else
	car_info = rs_car("car_owner") + "���� , ���� : " + rs_car("car_name") + " , ���� : " + rs_car("oil_kind")
	car_owner = rs_car("car_owner")
end if

end_yn = "N"
sql = "select * from cost_end where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"

set rs=dbconn.execute(sql)
if rs.eof or rs.bof then
	end_yn = "N"
else
	end_yn = rs("end_yn")
end if
rs.close()

'Response.write end_yn

' �����߰�
'sql = "call COST_PERSON_01('" & mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01" & "','"&user_id&"',@ret)"

   arParams = Array(mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", user_id)
   Set cmd = server.CreateObject("ADODB.Command")
   cmd.CommandText = query_person_cost
   Set cmd.ActiveConnection = dbconn
   Set rs = cmd.execute(,arParams,1)
'Response.write user_id
'   Response.write query_person_cost
'Response.write mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01"

'sql = "select * from person_cost where cost_month = '"&be_month&"' and emp_no = '"&user_id&"'"
'set rs=dbconn.execute(sql)
'Response.write sql
if rs.eof or rs.bof then
	'Response.write "AAAA"
	be_general_cnt = 0
	be_general_cost = 0
	be_overtime_cnt = 0
	be_overtime_cost = 0
	be_somopum_cost = 0
	be_fare_cnt = 0
	be_fare_cost = 0
	be_oil_cash_cost = 0
	be_repair_cost = 0
	be_parking_cost = 0
	be_toll_cost = 0
	be_card_cost = 0
	be_card_cost_vat = 0
	be_return_cash = 0
	be_tot_km = 0
	be_tot_cost = 0
	be_card_price = 0
	be_juyoo_card_price = 0
	be_cash_tot_cost = 0
else
	be_general_cnt = rs("general_cnt")
	be_general_cost = rs("general_cost")
	be_overtime_cnt = rs("overtime_cnt")
	be_overtime_cost = rs("overtime_cost")
	gas_km = cdbl(toString(rs("gas_km"),"0"))
	gas_unit = cdbl(toString(rs("gas_unit"),"0"))
	gas_cost = cdbl(toString(rs("gas_cost"),"0"))
	gasol_km = cdbl(toString(rs("gasol_km"),"0"))
	gasol_unit = cdbl(toString(rs("gasol_unit"),"0"))
	gasol_cost = cdbl(toString(rs("gasol_cost"),"0"))
	diesel_km = cdbl(toString(rs("diesel_km"),"0"))
	diesel_unit = cdbl(toString(rs("diesel_unit"),"0"))
	diesel_cost = cdbl(toString(rs("diesel_cost"),"0"))
	be_somopum_cost = cdbl(toString(rs("somopum_cost"),"0"))

	be_fare_cnt = rs("fare_cnt")
	be_fare_cost = rs("fare_cost")
	be_oil_cash_cost = rs("oil_cash_cost")
	be_repair_cost = rs("repair_cost")
	be_parking_cost = rs("parking_cost")
	be_toll_cost = rs("toll_cost")
	be_card_cost = rs("card_cost")
	be_card_cost_vat = rs("card_cost_vat")
	be_return_cash = rs("return_cash")

	be_tot_km = gas_km + diesel_km + gasol_km
	be_tot_cost = gas_cost + diesel_cost + gasol_cost
	be_card_price = be_card_cost + be_card_cost_vat
	be_juyoo_card_price = rs("juyoo_card_cost") + rs("juyoo_card_cost_vat")

	be_company_yn = cdbl(rs("company_yn"))
	if be_company_yn > 0 then
	  be_cash_tot_cost =  be_fare_cost + be_oil_cash_cost + be_toll_cost + be_parking_cost
	else
	   be_cash_tot_cost = be_general_cost + gas_cost + diesel_cost + gasol_cost + be_somopum_cost + be_fare_cost + be_oil_cash_cost + be_toll_cost + be_parking_cost
  end if
end if
rs.close()

if end_yn = "Y" then
'	response.write("read")
  'sql = "call COST_PERSON_01('" & mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01" & "','"&user_id&"',@ret)"
	'sql = "select * from person_cost where cost_month = '"&cost_month&"' and emp_no = '"&user_id&"'"
	'set rs=dbconn.execute(sql)

   arParams = Array(mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                    mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", user_id)
   Set cmd = server.CreateObject("ADODB.Command")
   cmd.CommandText = query_person_cost
   Set cmd.ActiveConnection = dbconn
   Set rs = cmd.execute(,arParams,1)

	general_cnt = rs("general_cnt")
	general_cost = rs("general_cost")
	overtime_cnt = rs("overtime_cnt")
	overtime_cost = rs("overtime_cost")
	gas_km = cdbl(toString(rs("gas_km"),"0"))
	gas_unit = cdbl(toString(rs("gas_unit"),"0"))
	gas_cost = cdbl(toString(rs("gas_cost"),"0"))
	gasol_km = cdbl(toString(rs("gasol_km"),"0"))
	gasol_unit = cdbl(toString(rs("gasol_unit"),"0"))
	gasol_cost = cdbl(toString(rs("gasol_cost"),"0"))
	diesel_km = cdbl(toString(rs("diesel_km"),"0"))
	diesel_unit = cdbl(toString(rs("diesel_unit"),"0"))
	diesel_cost = cdbl(toString(rs("diesel_cost"),"0"))
	somopum_cost = cdbl(toString(rs("somopum_cost"),"0"))
	fare_cnt = rs("fare_cnt")
	fare_cost = rs("fare_cost")
	oil_cash_cost = rs("oil_cash_cost")
	repair_cost = rs("repair_cost")
	parking_cost = rs("parking_cost")
	toll_cost = rs("toll_cost")
	card_cost = rs("card_cost")
	card_cost_vat = rs("card_cost_vat")
	return_cash = rs("return_cash")

	tot_km = gas_km + diesel_km + gasol_km
    tot_cost = gas_cost + diesel_cost + gasol_cost
	card_price = card_cost + card_cost_vat
	juyoo_card_price = rs("juyoo_card_cost") + rs("juyoo_card_cost_vat")
	'cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
	company_yn = cdbl(toString(rs("company_yn"),"0"))
  if company_yn > 0 then
    cash_tot_cost =   fare_cost + oil_cash_cost + toll_cost + parking_cost
  else
     cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
  end if

	variation_memo = rs("variation_memo")
		'response.write variation_memo
else
	sql = "select * from person_cost where cost_month = '"&cost_month&"' and emp_no = '"&user_id&"'"
	set rs=dbconn.execute(sql)
	if rs.eof or rs.bof then
		variation_memo = ""
	  else
		variation_memo = rs("variation_memo")
	end if
	rs.close()

	'response.write variation_memo
	' �Ϲݺ��
	general_cnt = 0
	general_cost = 0
	general_pre_cnt = 0
	general_pre_cost = 0
	sql = "select pay_yn,count(slip_seq) as c_cnt,sum(cost) as cost from general_cost where (emp_no='"&user_id&"') "& _
	"and (slip_gubun = '���') and (tax_bill_yn = 'N' or isnull(tax_bill_yn)) and (cancel_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by pay_yn"
	rs.Open sql, Dbconn, 1
	do until rs.eof
		if rs("pay_yn") = "N" then
			general_cnt = general_cnt + cint(rs("c_cnt"))
			general_cost = general_cost + cdbl(toString(rs("cost"),"0"))
		  else
			general_pre_cnt = general_pre_cnt + cint(rs("c_cnt"))
			general_pre_cost = general_pre_cost + cdbl(toString(rs("cost"),"0"))
		end if
		rs.movenext()
	loop
	rs.close()

	' ��Ư��
	overtime_cnt = 0
	overtime_cost = 0
	sql = "select cancel_yn,count(work_date) as c_cnt,sum(overtime_amt) as cost from overtime where (mg_ce_id='"&user_id&"') "& _
	"and (work_date >='"&from_date&"' and work_date <='"&to_date&"') and (cancel_yn = 'N') group by cancel_yn"

	rs.Open sql, Dbconn, 1
	do until rs.eof
		overtime_cnt = overtime_cnt + cint(rs("c_cnt"))
		overtime_cost = overtime_cost + cdbl(toString(rs("cost"),"0"))
		rs.movenext()
	loop
	rs.close()

	' �����
	gas_km = 0
	gas_km2 = 0
	gas_unit = 0
	gas_cost = 0
	diesel_km = 0
	diesel_km2 = 0
	diesel_unit = 0
	diesel_cost = 0
	gasol_km = 0
	gasol_km2 = 0
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
	company_yn =0
	sql = "select * from transit_cost where (mg_ce_id='"&user_id&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cancel_yn = 'N')"
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

	if rs("car_owner") = "ȸ��"  Then

      'if car_owner ="ȸ��" then ' �λ���� �ý��� /  �������� / ������Ȳ ���� �ش�����(������ȣ�� ��ȸ)�� ȸ��,���� ���θ� Ȯ���� ��... 19.01.14 �輺��
	      company_yn = company_yn + 1
'	      Response.write "2222"
'Response.write rs("car_owner")

		  if rs("oil_kind") = "�ֹ���" then
	        gasol_km2 = gasol_km2 + rs("far")
	      end if

		  if rs("oil_kind") = "����" then
	        diesel_km2 = diesel_km2 + rs("far")
	      end If

	      if rs("oil_kind") = "����" then
	        gas_km2 = gas_km2 + rs("far")
	      end If

	    'end if
    end If
'Response.write rs("oil_kind")&"<br>"
'Response.write rs("far")&"<br>"

		if rs("car_owner") = "ȸ��" then
			oil_cash_cost = oil_cash_cost + rs("oil_price")
			repair_cost = repair_cost + rs("repair_cost")
		end if

		parking_cost = parking_cost + rs("parking")
		toll_cost = toll_cost + rs("toll")
		rs.movenext()
	loop
	rs.close()

	if team = "������" or team = "SM1��" or team = "Repair��" or team = "SM2��" then
		oil_unit_id = "1"
	  else
		oil_unit_id = "2"
	end If

    sql = "select * from oil_unit where oil_unit_month = '"&cost_month&"' and oil_unit_id = '"&oil_unit_id&"'"
'Response.write sql&"<br>"
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

	if reside_company = "��ȭȭ��" then
		liter = 8
	  else
		liter = 10
	end if

	tot_km = gas_km + diesel_km + gasol_km
'Response.Write "gas_km"&gas_km&"<br>"
'Response.Write "diesel_km"&diesel_km&"<br>"
'Response.Write "gasol_km"&gasol_km&"<br>"
	somopum_cost = (tot_km) * 25

	gas_cost = round(gas_km * gas_unit / 7)
	diesel_cost = round(diesel_km * diesel_unit / liter)
	gasol_cost = round(gasol_km * gasol_unit / liter)
	tot_cost = gas_cost + diesel_cost + gasol_cost
'Response.Write "gas_unit"&gas_unit&"<br>"
'Response.Write "diesel_km"&diesel_km&"<br>"
'Response.Write "gasol_km"&gasol_unit&"<br>"

'Response.Write "gas_cost"&gas_cost&"<br>"
'Response.Write "diesel_cost"&diesel_cost&"<br>"
'Response.Write "gasol_cost"&gasol_cost&"<br>"


		' ���� ī����
		juyoo_card_cnt = 0
		juyoo_card_cost = 0
		juyoo_card_cost_vat = 0
		juyoo_card_price = 0
		sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&user_id&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and card_type like '%����%'"

		Set rs = Dbconn.Execute (sql)
		if cint(rs("c_cnt")) <>  0 then
			juyoo_card_cnt = juyoo_card_cnt + cint(rs("c_cnt"))
			juyoo_card_cost = juyoo_card_cost + cdbl(toString(rs("cost"),"0"))
			juyoo_card_cost_vat = juyoo_card_cost_vat + cdbl(toString(rs("cost_vat"),"0"))
		end if
		rs.close()
		juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat

	' ī����
	card_cnt = 0
	card_cost = 0
	card_cost_vat = 0
	card_price = 0
	sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&user_id&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and card_type not like '%����%'"
'	sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&user_id&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"

	Set rs = Dbconn.Execute (sql)
'	Response.write sql
	if cint(rs("c_cnt")) <>  0 then
		card_cnt = card_cnt + cint(rs("c_cnt"))
		card_cost = card_cost + cdbl(toString(rs("cost"),"0"))
		card_cost_vat = card_cost_vat + cdbl(toString(rs("cost_vat"),"0"))
	end if
	rs.close()
	card_price = card_cost + card_cost_vat

	cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost

'	if tot_km <> 0 then
	if car_owner = "����" then
		return_cash = cash_tot_cost - juyoo_card_price
	else
		return_cash = cash_tot_cost
	end if


	dbconn.BeginTrans
	sql = "delete from person_cost where cost_month ='"&cost_month&"' and emp_no ='"&user_id&"'"
	dbconn.execute(sql)

	sql = "insert into person_cost values ('"&cost_month&"','"&user_id&"','"&user_name&"','"&user_grade&"','�ٹ�','"&car_owner&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&reside_company&"',"&general_cnt&","&general_cost&","&general_pre_cnt&","&general_pre_cost&","&overtime_cnt&","&overtime_cost&","&gas_km&","&gas_unit&","&gas_cost&","&diesel_km&","&diesel_unit&","&diesel_cost&","&gasol_km&","&gasol_unit&","&gasol_cost&","&tot_km&","&tot_cost&","&somopum_cost&","&fare_cnt&","&fare_cost&","&oil_cash_cost&","&repair_cost&","&repair_pre_cost&","&parking_cost&","&toll_cost&","&juyoo_card_cnt&","&juyoo_card_cost&","&juyoo_card_cost_vat&","&card_cnt&","&card_cost&","&card_cost_vat&","&return_cash&",'"&variation_memo&"',now(), 0)"
	dbconn.execute(sql)



	if(company_yn > 0) then
   tot_km = gas_km2 + diesel_km2 + gasol_km2
   somopum_cost = (tot_km) * 25

   gas_cost = round(gas_km2 * gas_unit / 7)
   diesel_cost = round(diesel_km2 * diesel_unit / liter)
   gasol_cost = round(gasol_km2 * gasol_unit / liter)
   tot_cost = gas_cost + diesel_cost + gasol_cost
	end if
  if car_owner="����" then
     car_owner =""
  end if
	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "ó���� Error�� �߻��Ͽ����ϴ�...."
		response.write"<script language=javascript>"
		response.write"alert('"&end_msg&"');"
		response.write"location.replace('person_cost_report.asp');"
		response.write"</script>"
		Response.End
	  else
		dbconn.CommitTrans
	end if
end if
i = 1
title_line = "���κ� ��� ���� ��Ȳ"
'Response.write sql
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("�߻������ �Է��ϼ���.");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="person_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�߻����&nbsp;</strong>(��201401) :
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px" maxlength="6">
								</label>
								<label>
								<strong>�������� : </strong><%=car_info%>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
				<h3 class="stit">* ������ ��� Ȯ���� ���κ� ��� ���� ��ǥ ����� �����ڷḦ ÷���Ͻø� �˴ϴ�. (��ǥ����� ���������� �Է��ؾ� ����� �����մϴ�.)</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="4%" >
							<col width="6%" >
							<col width="4%" >
							<col width="6%" >
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="3" class="first" scope="col">���</th>
								<th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">��Ư��</th>
								<th colspan="11" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� ���</th>
								<th rowspan="3" scope="col">����ī��</th>
								<th rowspan="3" scope="col">����ī��<br>(VAT����)</th>
								<th rowspan="3" scope="col">����ݾ�</th>
							</tr>
							<tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;border-left:1px solid #e3e3e3;" scope="col">��û�ݾ�</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�Ϲݺ��</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">���߱����</th>
							  <th colspan="3" style=" border-bottom:1px solid #e3e3e3;" scope="col"><%=car_owner%> ���� ������</th>
							  <th style=" border-bottom:1px solid #e3e3e3;" scope="col">ȸ������</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">���� ������</th>
							  <th rowspan="2" scope="col"><p>���ݻ��</p><p>�Ұ�</p></th>
						  </tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">�Ǽ�</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">�Ǽ�</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">�Ǽ�</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">����(KM)</th>
							  <th scope="col">������</th>
							  <th scope="col">�Ҹ�ǰ</th>
							  <th scope="col">������</th>
							  <th scope="col">������</th>
							  <th scope="col">�����</th>
						  </tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFFFCC">
								<td height="25" class="first"><strong>����</strong></td>
								<td class="right"><%=formatnumber(be_overtime_cnt,0)%></td>
								<td class="right">0</td><!-- <%=formatnumber(be_overtime_cost,0)%> --> <!--180917 ������ �䱸-->
								<td class="right"><%=formatnumber(be_general_cnt,0)%></td>
								<td class="right"><%=formatnumber(be_general_cost,0)%></td>
								<td class="right"><%=formatnumber(be_fare_cnt,0)%></td>
								<td class="right"><%=formatnumber(be_fare_cost,0)%></td>
								<td class="right"><%=formatnumber(be_tot_km,0)%></td>
								<td class="right"><%=formatnumber(be_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(be_somopum_cost,0)%></td>
								<td class="right"><%=formatnumber(be_oil_cash_cost,0)%></td>
								<td class="right"><%=formatnumber(be_parking_cost,0)%></td>
								<td class="right"><%=formatnumber(be_toll_cost,0)%></td>
								<td class="right"><%=formatnumber(be_cash_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(be_juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(be_card_cost,0)%></td>
								<td class="right"><%=formatnumber(be_return_cash,0)%></td>
							</tr>
							<tr bgcolor="#E8FFFF">
								<td height="25" class="first"><strong>���</strong></td>
			      				<td class="right"><%=formatnumber(overtime_cnt,0)%></td>
								<td class="right">0</td><!-- <%=formatnumber(overtime_cost,0)%> --> <!--180917 ������ �䱸-->
								<td class="right"><%=formatnumber(general_cnt,0)%></td>
								<td class="right"><%=formatnumber(general_cost,0)%></td>
								<td class="right"><%=formatnumber(fare_cnt,0)%></td>
								<td class="right"><%=formatnumber(fare_cost,0)%></td>
								<td class="right"><%=formatnumber(tot_km,0)%></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum_cost,0)%></td>
								<td class="right"><%=formatnumber(oil_cash_cost,0)%></td>
								<td class="right"><%=formatnumber(parking_cost,0)%></td>
								<td class="right"><%=formatnumber(toll_cost,0)%></td>
								<td class="right"><%=formatnumber(cash_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(card_cost,0)%></td>
								<td class="right"><%=formatnumber(return_cash,0)%></td>
							</tr>
					<%
					overtime_cal = overtime_cost - be_overtime_cost
					if be_overtime_cost = 0 then
						overtime_per = 100
					end if
					if overtime_cost = 0 then
						overtime_per = -100
					end if
					if overtime_cost = 0 and be_overtime_cost = 0 then
						overtime_per = 0
					end if
					if overtime_cost <> 0 and be_overtime_cost <> 0 then
						overtime_per = overtime_cal / be_overtime_cost * 100
					end if

					general_cal = general_cost - be_general_cost
					if be_general_cost = 0 then
						general_per = 100
					end if
					if general_cost = 0 then
						general_per = -100
					end if
					if general_cost = 0 and be_general_cost = 0 then
						general_per = 0
					end if
					if general_cost <> 0 and be_general_cost <> 0 then
						general_per = general_cal / be_general_cost * 100
					end if

					fare_cal = fare_cost - be_fare_cost
					if be_fare_cost = 0 then
						fare_per = 100
					end if
					if fare_cost = 0 then
						fare_per = -100
					end if
					if fare_cost = 0 and be_fare_cost = 0 then
						fare_per = 0
					end if
					if fare_cost <> 0 and be_fare_cost <> 0 then
						fare_per = fare_cal / be_fare_cost * 100
					end if

					tot_km_cal = tot_km - be_tot_km
					if be_tot_km = 0 then
						tot_km_per = 100
					end if
					if tot_km = 0 then
						tot_km_per = -100
					end if
					if tot_km = 0 and be_tot_km = 0 then
						tot_km_per = 0
					end if
					if tot_km <> 0 and be_tot_km <> 0 then
						tot_km_per = tot_km_cal / be_tot_km * 100
					end if

					tot_cost_cal = tot_cost - be_tot_cost
					if be_tot_cost = 0 then
						tot_cost_per = 100
					end if
					if tot_cost = 0 then
						tot_cost_per = -100
					end if
					if tot_cost = 0 and be_tot_cost = 0 then
						tot_cost_per = 0
					end if
					if tot_cost <> 0 and be_tot_cost <> 0 then
						tot_cost_per = tot_cost_cal / be_tot_cost * 100
					end if

					somopum_cost_cal = somopum_cost - be_somopum_cost
					if be_somopum_cost = 0 then
						somopum_per = 100
					end if
					if somopum_cost = 0 then
						somopum_per = -100
					end if
					if somopum_cost = 0 and be_somopum_cost = 0 then
						somopum_per = 0
					end if
					if somopum_cost <> 0 and be_somopum_cost <> 0 then
						somopum_per = somopum_cost_cal / be_somopum_cost * 100
					end if

					oil_cash_cost_cal = oil_cash_cost - be_oil_cash_cost
					if be_oil_cash_cost = 0 then
						oil_cash_per = 100
					end if
					if oil_cash_cost = 0 then
						oil_cash_per = -100
					end if
					if oil_cash_cost = 0 and be_oil_cash_cost = 0 then
						oil_cash_per = 0
					end if
					if oil_cash_cost <> 0 and be_oil_cash_cost <> 0 then
						oil_cash_per = oil_cash_cost_cal / be_oil_cash_cost * 100
					end if

					parking_cost_cal = parking_cost - be_parking_cost
					if be_parking_cost = 0 then
						parking_per = 100
					end if
					if parking_cost = 0 then
						parking_per = -100
					end if
					if parking_cost = 0 and be_parking_cost = 0 then
						parking_per = 0
					end if
					if parking_cost <> 0 and be_parking_cost <> 0 then
						parking_per = parking_cost_cal / be_parking_cost * 100
					end if

					cash_tot_cost_cal = cash_tot_cost - be_cash_tot_cost
					if be_cash_tot_cost = 0 then
						cash_tot_per = 100
					end if
					if cash_tot_cost = 0 then
						cash_tot_per = -100
					end if
					if cash_tot_cost = 0 and be_cash_tot_cost = 0 then
						cash_tot_per = 0
					end if
					if cash_tot_cost <> 0 and be_cash_tot_cost <> 0 then
						cash_tot_per = cash_tot_cost_cal / be_cash_tot_cost * 100
					end if

					juyoo_card_price_cal = juyoo_card_price - be_juyoo_card_price
					if be_juyoo_card_price = 0 then
						juyoo_card_per = 100
					end if
					if juyoo_card_price = 0 then
						juyoo_card_per = -100
					end if
					if juyoo_card_price = 0 and be_juyoo_card_price = 0 then
						juyoo_card_per = 0
					end if
					if juyoo_card_price <> 0 and be_juyoo_card_price <> 0 then
						juyoo_card_per = juyoo_card_price_cal / be_juyoo_card_price * 100
					end if

					card_cost_cal = card_cost - be_card_cost
					if be_card_cost = 0 then
						card_per = 100
					end if
					if card_cost = 0 then
						card_per = -100
					end if
					if card_cost = 0 and be_card_cost = 0 then
						card_per = 0
					end if
					if card_cost <> 0 and be_card_cost <> 0 then
						card_per = card_cost_cal / be_card_cost * 100
					end if

					return_cash_cal = return_cash - be_return_cash
					if be_return_cash = 0 then
						return_per = 100
					end if
					if return_cash = 0 then
						return_per = -100
					end if
					if return_cash = 0 and be_return_cash = 0 then
						return_per = 0
					end if
					if return_cash <> 0 and be_return_cash <> 0 then
						return_per = return_cash_cal / be_return_cash * 100
					end if

					%>
							<tr bgcolor="#FFE8E8">
								<td height="25" class="first"><strong>����(%)</strong></td>
				      	<td class="right">&nbsp;</td>
								<td class="right">0%</td><!-- <%=formatnumber(overtime_per,2)%> --> <!--180917 ������ �䱸-->
								<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(general_per,2)%>%</td>
								<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(fare_per,2)%>%</td>
								<td class="right"><%=formatnumber(tot_km_per,2)%>%</td>
								<td class="right"><%=formatnumber(tot_cost_per,2)%>%</td>
								<td class="right"><%=formatnumber(somopum_per,2)%>%</td>
								<td class="right"><%=formatnumber(oil_cash_per,2)%>%</td>
								<td class="right"><%=formatnumber(parking_per,2)%>%</td>
								<td class="right"><%=formatnumber(toll_per,2)%>%</td>
								<td class="right"><%=formatnumber(cash_tot_per,2)%>%</td>
								<td class="right"><%=formatnumber(juyoo_card_per,2)%>%</td>
								<td class="right"><%=formatnumber(card_per,2)%>%</td>
								<td class="right"><%=formatnumber(return_per,2)%>%</td>
						  </tr>
							<tr>
								<td height="25" class="first"><strong>��������</strong></td>
				      		  	<td colspan="16" class="left"><%=variation_memo%>&nbsp;</td>
						  </tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
				<% if end_yn <> "Y" then	%>
					<a href="#" onClick="pop_Window('variation_memo_add.asp?cost_month=<%=cost_month%>&emp_no=<%=user_id%>','variation_memo_add_pop','scrollbars=yes,width=600,height=300')" class="btnType04">����������Ϲ׺���</a>
				<% end if	%>
					</div>
                  	</td>
				    <td width="50%">
                  	</td>
				    <td width="25%">
					<div class="btnCenter">
				<% if variation_memo <> "" then	%>
                    <a href="#" onClick="pop_Window('person_cost_print.asp?cost_month=<%=cost_month%>','person_cost_print_popup','scrollbars=yes,width=1250,height=650')" class="btnType04">���κ� ��� ���� ��ǥ���</a>
				<% end if %>
					</div>
                    </td>
			      </tr>
				  </table>
				<br>
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
			</form>
		</div>
	</div>
	</body>
</html>
