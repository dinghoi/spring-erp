<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

'on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

inc_yyyy = cint(mid(now(),1,4)) - 1

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
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
Set rs_othe = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")
emp_person = cstr(rs_emp("emp_person1")) + "-" + cstr(rs_emp("emp_person2"))
rs_emp.close()

tot_pay = 0
y_nhis_amt = 0
y_longcare_amt = 0
y_epi_amt = 0

y_nps_other = 0
y_nps_amt = 0

Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
if not rs_year.eof then
       y_nps_amt = rs_year("y_nps_amt")
	   y_nhis_amt = rs_year("y_nhis_amt")
	   y_longcare_amt = rs_year("y_longcare_amt")
	   y_epi_amt = rs_year("y_epi_amt")
	   tot_pay = rs_year("y_total_pay") + rs_year("y_total_bonus") + rs_year("y_other_pay")
	   y_householder = rs_year("y_householder")
	   y_national = rs_year("y_national")
	   y_from_date = rs_year("y_from_date")
	   y_to_date = rs_year("y_to_date")
	   y_live = rs_year("y_live")
	   y_change = rs_year("y_change")
	   y_basic_cnt =  rs_year("y_wife") + rs_year("y_age20_cnt") + rs_year("y_age60_cnt")
       bon_person_no =  cstr(rs_year("y_person_no1")) + cstr(rs_year("y_person_no2"))

	if rs_year("y_woman") = "Y" then
		y_woman = "��"
	else
		y_woman = ""
	end If

	if rs_year("y_single") = "Y" then
		y_single = "��"
	else
		y_single = ""
	end if
else
       y_nps_amt = 0
	   y_nhis_amt = 0
	   y_longcare_amt = 0
	   y_epi_amt = 0
	   y_householder = "N"
	   y_national = ""
	   y_live = "Y"
	   y_change = "N"

	   y_woman = ""
	   y_single = ""
end if

y_nps_tax = y_nps_amt
y_nhis_amt = y_nhis_amt + y_longcare_amt
y_nhis_tax = y_nhis_amt
y_epi_tax = y_epi_amt

'if rs_year("y_woman") = "Y" then
'	y_woman = "��"
'else
'	y_woman = ""
'end If

'if rs_year("y_single") = "Y" then
'	y_single = "��"
'else
'	y_single = ""
'end if

y_national_code = "001"

if y_householder = "Y" then
       householder = "[��]������ [ ]�����"
   else
       householder = "[ ]������ [��]�����"
end if

if y_live = "Y" then
       yy_live = "[��]������ [ ]�������"
   else
       yy_live = "[ ]������ [��]�������"
end if

if y_change = "N" then
       yy_change = "[��]��������� [ ]����"
   else
       yy_change = "[ ]��������� [��]����"
end if

'�ſ�ī��
sum_nts_market = 0
sum_nts_transit = 0
sum_nts_credit = 0
sum_cash = 0
sum_nts_direct = 0
sum_nts_donation = 0
sum_nts_edu = 0
sum_nts_medical = 0
sum_nts_insuran = 0

sum_oth_market = 0
sum_oth_transit = 0
sum_oth_credit = 0
sum_oth_direct = 0
sum_oth_donation = 0
sum_oth_edu = 0
sum_oth_medical = 0
sum_oth_insuran = 0
sum_oth_nhis = 0
sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
do until rs_fami.eof
         sum_nts_credit = sum_nts_credit + rs_fami("c_credit_nts")
		 sum_oth_credit = sum_oth_credit + rs_fami("c_credit_other")
		 sum_nts_direct = sum_nts_direct + rs_fami("c_direct_nts")
		 sum_oth_direct = sum_oth_direct + rs_fami("c_direct_other")
		 sum_cash = sum_cash + rs_fami("c_cash_nts")
		 sum_nts_market = sum_nts_market + rs_fami("c_market_nts")
		 sum_oth_market = sum_oth_market + rs_fami("c_market_other")
		 sum_nts_transit = sum_nts_transit + rs_fami("c_transit_nts")
		 sum_oth_transit = sum_oth_transit + rs_fami("c_transit_other")

		 sum_nts_donation = sum_nts_donation + rs_fami("d_poli_nts") + rs_fami("d_poli10_nts") + rs_fami("d_law_nts") + rs_fami("d_ji_nts")
		 sum_oth_donation = sum_oth_donation + rs_fami("d_poli_other") + rs_fami("d_poli10_other") + rs_fami("d_law_other") + rs_fami("d_ji_other")

		 sum_nts_edu = sum_nts_edu + rs_fami("e_nts_amt")
		 sum_oth_edu = sum_oth_edu + rs_fami("e_other_amt")

		 sum_nts_medical = sum_nts_medical + rs_fami("m_nts_amt")
		 sum_oth_medical = sum_oth_medical + rs_fami("m_other_amt")

		 sum_nts_insuran = sum_nts_insuran + rs_fami("i_ilban_nts") + rs_fami("i_disab_nts")
		 sum_oth_insuran = sum_oth_insuran + rs_fami("i_ilban_other") + rs_fami("i_disab_other")
	rs_fami.MoveNext()
loop
rs_fami.close()

'sum_nts_nhis =  rs_year("y_nhis_amt") + rs_year("y_epi_amt") + rs_year("y_longcare_amt")
sum_nts_nhis =  y_nhis_amt + y_epi_amt + y_longcare_amt

b_nps = 0
b_nhis = 0
b_longcare = 0
b_epi = 0
Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
'Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
	   tot_pay = tot_pay + rs_bef("b_pay") + rs_bef("b_bonus") + rs_bef("b_deem_bonus")
	   b_nps = b_nps + rs_bef("b_nps")
	   b_nhis = b_nhis + rs_bef("b_nhis")
	   b_longcare = b_longcare + rs_bef("b_longcare")
	   b_epi = b_epi + rs_bef("b_epi")
	rs_bef.MoveNext()
loop
rs_bef.close()

b_nps_tax = b_nps
b_nhis = b_nhis + b_longcare
b_nhis_tax = b_nhis
b_epi_tax = b_epi

'���ݺ���
a_amt_other = 0
a_amt_tot = 0
Sql = "select * from pay_yeartax_annuity where a_year = '"&inc_yyyy&"' and a_emp_no = '"&emp_no&"' ORDER BY a_emp_no,a_seq ASC"
rs_ann.Open Sql, Dbconn, 1
Set rs_ann = DbConn.Execute(SQL)
do until rs_ann.eof
       a_amt_tot = a_amt_tot + rs_ann("a_amt")
	rs_ann.MoveNext()
loop
rs_ann.close()

a_amt_tax = a_amt_tot

a_amt_other_tax = a_amt_other
y_nps_other_tax = y_nps_other

a_tot_amt = y_nps_amt + b_nps + y_nps_other
a_tot_tax = y_nps_tax + b_nps_tax + y_nps_other_tax

'�����
ilban_insu = 0
disab_insu = 0
Sql = "select * from pay_yeartax_insurance where i_year = '"&inc_yyyy&"' and i_emp_no = '"&emp_no&"' ORDER BY i_emp_no,i_seq ASC"
rs_ins.Open Sql, Dbconn, 1
Set rs_ins = DbConn.Execute(SQL)
do until rs_ins.eof
       if rs_ins("i_disab_chk") = "Y" then
	          disab_insu = disab_insu + rs_ins("i_nts_amt") + rs_ins("i_other_amt")
		  else
			  ilban_insu = ilban_insu + rs_ins("i_nts_amt") + rs_ins("i_other_amt")
	   end if
	rs_ins.MoveNext()
loop
rs_ins.close()

if ilban_insu > 1000000 then
       ilban_insu_tax = 1000000
   else
       ilban_insu_tax = ilban_insu
end if

if disab_insu > 1000000 then
       disab_insu_tax = 1000000
   else
       disab_insu_tax = disab_insu
end if

i_tot_amt = y_nhis_amt + b_nhis + y_epi_amt + b_epi
i_tot_tax = y_nhis_tax + b_nhis_tax + y_epi_tax + b_epi_tax

'�����ڱ�
Sql = "select * from pay_yeartax_house where h_year = '"&inc_yyyy&"' and h_emp_no = '"&emp_no&"'"
rs_hous.Open Sql, Dbconn, 1
Set rs_hous = DbConn.Execute(SQL)
if not rs_hous.eof then
       u_type = "U"
       h_lender_amt = rs_hous("h_lender_amt")
	   h_person_amt = rs_hous("h_person_amt")
	   h_long15_amt = rs_hous("h_long15_amt")
	   h_long29_amt = rs_hous("h_long29_amt")
	   h_long30_amt = rs_hous("h_long30_amt")
	   h_fixed_amt = rs_hous("h_fixed_amt")
	   h_other_amt = rs_hous("h_other_amt")
   else
       u_type = ""
       h_lender_amt = 0
	   h_person_amt = 0
	   h_long15_amt = 0
	   h_long29_amt = 0
	   h_long30_amt = 0
	   h_fixed_amt = 0
	   h_other_amt = 0
end if
rs_hous.close()

h_tot_amt = h_lender_amt + h_person_amt + h_long15_amt + h_long29_amt + h_long30_amt + h_fixed_amt + h_other_amt
h_tot_tax = h_lender_tax + h_person_tax + h_long15_tax + h_long29_tax + h_long30_tax + h_fixed_tax + h_other_tax

h_month_amt = 0
Sql = "select * from pay_yeartax_house_m where hm_year = '"&inc_yyyy&"' and hm_emp_no = '"&emp_no&"' ORDER BY hm_emp_no,hm_seq ASC"
rs_houm.Open Sql, Dbconn, 1
Set rs_houm = DbConn.Execute(SQL)
do until rs_houm.eof
       h_month_amt = h_month_amt + rs_houm("hm_month_amt")
	rs_houm.MoveNext()
loop
rs_houm.close()

'��������
s_id = "��������"
tot_2000 = 0
tot_2001 = 0
tot_endi = 0
Sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
rs_savi.Open Sql, Dbconn, 1
Set rs_savi = DbConn.Execute(SQL)
do until rs_savi.eof
       if rs_savi("s_type") = "���ο�������(2000������)" then
	           tot_2000 = tot_2000 + rs_savi("s_amt")
		  elseif rs_savi("s_type") = "��������(2001������)" then
	                  tot_2001 = tot_2001 + rs_savi("s_amt")
			     elseif rs_savi("s_type") = "�������ݼҵ����" then
	                         tot_endi = tot_endi + rs_savi("s_amt")
		end if
	rs_savi.MoveNext()
loop
rs_savi.close()

tax_2000 = tot_2000
tax_2001 = tot_2001
tax_endi = tot_endi

oy_tot_amt = tot_2000 + tot_2001 + tot_endi
oy_tot_tax = tax_2000 + tax_2001 + tax_endi

'��Ÿ����/������������
Sql = "select * from pay_yeartax_other where o_year = '"&inc_yyyy&"' and o_emp_no = '"&emp_no&"'"
rs_othe.Open Sql, Dbconn, 1
Set rs_othe = DbConn.Execute(SQL)
if not rs_othe.eof then
       u_type = "U"
       o_nps = rs_othe("o_nps")
	   o_nhis = rs_othe("o_nhis")
	   o_sosang = rs_othe("o_sosang")
	   o_chul2012 = rs_othe("o_chul2012")
	   o_chul2013 = rs_othe("o_chul2013")
	   o_chul2014 = rs_othe("o_chul2014")
	   o_woori = rs_othe("o_woori")
	   o_goyoung = rs_othe("o_goyoung")
	   o_chul_hap = o_chul2008 + o_chul2009
   else
       u_type = ""
       o_nps = 0
	   o_nhis = 0
	   o_sosang = 0
	   o_chul2012 = 0
	   o_chul2013 = 0
	   o_chul2014 = 0
	   o_woori = 0
	   o_goyoung = 0
	   o_chul_hap = 0
end if

 s_id = "���ø�������"
      tot_cheng = 0
      tot_jutak = 0
      tot_gunro = 0
	  tot_jangi = 0
      Sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
      rs_savi.Open Sql, Dbconn, 1
      Set rs_savi = DbConn.Execute(SQL)
      do until rs_savi.eof
            if rs_savi("s_type") = "û������" then
	                 tot_cheng = tot_cheng + rs_savi("s_amt")
		       elseif rs_savi("s_type") = "����û����������" then
	                        tot_jutak = tot_jutak + rs_savi("s_amt")
			          elseif rs_savi("s_type") = "�ٷ������ø�������" then
	                              tot_gunro = tot_gunro + rs_savi("s_amt")
							 elseif rs_savi("s_type") = "������ø�������" then
	                                 tot_jangi = tot_jangi + rs_savi("s_amt")
		    end if
	        rs_savi.MoveNext()
      loop
      rs_savi.close()

      tax_cheng = tot_cheng
      tax_jutak = tot_jutak
      tax_gunro = tot_gunro
	  tax_jangi = tot_jangi

      oj_tot_amt = tot_cheng + tot_jutak + tot_gunro
      oj_tot_tax = tax_cheng + tax_jutak + tax_gunro






rs_othe.close()
if emp_company = "���̿��������" then
      company_name = "(��)" + "���̿��������"
	  owner_name = "�����"
	  addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif emp_company = "�޵�" then
              company_name = "(��)" + "�޵�"
			  owner_name = "������"
	          addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif emp_company = "���̳�Ʈ����" then
                     company_name = "���̳�Ʈ����" + "(��)"
					 owner_name = "���߿�"
	                 addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif emp_company = "����������ġ" then
                        company_name = "(��)" + "����������ġ"
						owner_name = "�ڹ̾�"
	                    addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if

c_hap1 = 0

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�ҵ桤���� �����Ű�/�ٷ��� �ҵ桤���� �����Ű�(2014�� �ҵ濡 ���� ���������)"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.inc_yyyy.value == "") {
					alert ("�ͼӳ⵵�� �Է��ϼ���.");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_tax_report.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
							<col width="20%" >
							<col width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td colspan="7" style="font-size:14px;"><%=title_line%></td>
						    </tr>
                            <tr>
                              <td colspan="7" class="left">�� �ٷμҵ��ڴ� �Ű��� �ҵ桤���� ���� ������ ÷���Ͽ� ��õ¡���ǹ���(�Ҽ�ȸ�� ��)���� �����ϸ�, ��õ¡���ǹ��ڴ� �Ű� �� ÷�μ����� Ȯ���Ͽ� �ٷμҵ� ���װ���� �ϰ� �ٷμҵ��ڿ��� ���<br>�ٷμҵ��õ¡���������� �߱��ؾ� �մϴ�. �������� �� �ٷμҵ��ڿ��� ȯ���� �߻��ϴ� ��� ��õ¡���ǹ��ڴ� �ٷμҵ��ڿ��� ȯ�޼����� �����ؾ� �մϴ�.</td>
						    </tr>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">����</th>
                              <td style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%></td>
                              <th style=" border-top:1px solid #e3e3e3;">�λ��ڵ�</th>
                              <td colspan="2"><%=emp_no%></td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">�ֹε�Ϲ�ȣ(�Ǵ� �ܱ��ε�Ϲ�ȣ)</th>
                              <td ><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">���θ�</th>
                              <td colspan="3"><%=company_name%></td>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">��ü��</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">������ ����</th>
                              <td colspan="3"><%=householder%></td>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td>(�����ڵ�:<%=y_national_code%> )&nbsp;<%=y_national%></td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">�ٹ��Ⱓ</th>
                              <td colspan="3"><%=y_from_date%>&nbsp;��&nbsp;<%=y_to_date%></td>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">����Ⱓ</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">���ֱ���</th>
                              <td colspan="3"><%=yy_live%></td>
                              <th class="left">��������</th>
                              <td><%=y_national%>&nbsp;(�����ڵ�:<%=y_national_code%>)</td>
						    </tr>
                            <tr>
							  <th colspan="2">�������� �׸� ���� ����</th>
                              <td colspan="3"><%=yy_change%></td>
                              <td colspan="2" class="left" style="color:#ff0000;">�� �������� �׸��� ����� ������ ��쿡�� �ֹε��ǥ��� �������ֽñ� �ٶ��ϴ�.</td>
						    </tr>
                            <tr>
                              <td colspan="7">&nbsp;</td>
						    </tr>
						</thead>
					</table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
							<col width="4%" >
                            <col width="3%" >
							<col width="8%" >
							<col width="2%" >
                            <col width="2%" >
							<col width="3%" >
                            <col width="3%" >
                            <col width="7%" >

                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            </colgroup>
                            <thead>
                            <tr>
								<td rowspan="3" style=" border-right:1px solid #e3e3e3; border-bottom:1px solid #FFFFFF;">&nbsp;</td>
                                <th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3;">���������׸�</th>
                                <th colspan="11" scope="col" style=" border-bottom:1px solid #e3e3e3;">���� �ҵ���� �׸�</th>
							</tr>
                            <tr>
                                <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�⺻<br>����</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">���<br>���</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">���<br>�Ծ�</th>
                                <th rowspan="2" scope="col">�ڷᱸ��</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                                <th rowspan="2" scope="col">�Ƿ��</th>
                                <th rowspan="2" scope="col">������</th>
                                <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�ſ�ī�� �� ����</th>
                                <th rowspan="2" scope="col">��α�</th>
							</tr>
                            <tr>
                                <th class="first" scope="col">����</th>
                                <th scope="col">�ֹε�Ϲ�ȣ</th>
								<th scope="col">�γ���</th>
                                <th scope="col">�Ѻθ�</th>
                                <th scope="col">�����</th>
                                <th scope="col">6������</th>
                                <th scope="col">�ǰ�.����</th>
                                <th scope="col">���强</th>
                                <th scope="col">�ſ�ī��<br>(������塤���߱�������)</th>
                                <th scope="col">����ī�� ��<br>(������塤���߱�������)</th>
                                <th scope="col">���ݿ�����<br>(������塤���߱�������)</th>
                                <th scope="col">�������<br>����</th>
                                <th scope="col">���߱���<br>�̿��</th>
							</tr>
                            </thead>
                            <tbody>
                            <tr>
                                <td rowspan="20" style=" border-right:1px solid #e3e3e3;">��. �������� �� �ҵ���� ����</td>
                                <td rowspan="2" colspan="2" style=" border-top:2px solid #515254;">�ο� :&nbsp;<%=rs_year("y_support_cnt")%>&nbsp;(���ڳ�:&nbsp;<%=rs_year("y_daja_cnt")%>��)</td>
                                <td colspan="2" style=" border-top:2px solid #515254;"><%=y_basic_cnt%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=rs_year("y_old_cnt")%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=rs_year("y_holt_cnt")%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;">����û�ڷ�</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_nhis,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_insuran,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_medical,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_edu,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_credit,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_direct,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_cash,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_market,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_transit,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(sum_nts_donation,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;"><%=y_woman%>&nbsp;</td>
                                <td><%=y_single%>&nbsp;</td>
                                <td><%=rs_year("y_disab_cnt")%>&nbsp;</td>
                                <td><%=rs_year("y_age6_cnt")%>&nbsp;</td>
                                <td>�׹����ڷ�</td>
                                <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_insuran,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_medical,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_edu,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_credit,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_direct,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_transit,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_oth_donation,0)%>&nbsp;</td>
							</tr>
						<%
						do until rs.eof
                            f_rel = rs("f_rel")
							f_pensioner = rs("f_pensioner")
							f_witak = rs("f_witak")
							rel_chk = ""
							if f_rel = "����" then
							      f_person_no = "�ٷ��� ����"
								  rel_chk = "0"
							   else
							      f_person_no = cstr(mid(rs("f_person_no"),1,6)) + "-" + cstr(mid(rs("f_person_no"),7,7))
								  if f_rel = "��" or f_rel = "��" or f_rel = "����" or f_rel = "����" then
								         rel_chk = "1"
								     elseif f_rel = "����" or f_rel = "���" or f_rel = "������" or f_rel = "������" then
									            rel_chk = "2"
											elseif f_rel = "����" or f_rel = "�Ƴ�" then
											           rel_chk = "3"
											elseif f_rel = "�Ƶ�" or f_rel = "��" then
											           rel_chk = "4"
												   else
												       rel_chk = "5"
								  end if
						    end if
							if f_rel = "��(�����ڸ�)" or f_rel = "��(�����ڸ�)" or f_rel = "��(�����ڸ�)" or f_rel = "��(�����ڸ�)" then
							   rel_chk = "6"
							end if
							if f_pensioner = "Y" then
							   rel_chk = "7"
							end if
							if f_witak = "Y" then
							   rel_chk = "8"
							end if
							if rs("f_rel") = "����" or rs("f_wife") = "Y" or rs("f_age20") = "Y" or rs("f_age60") = "Y" or rs("f_old") = "Y" then
							        basic_chk = "��"
							   else
							        basic_chk = ""
						    end if
							if rs("f_old") = "Y" then
							        old_chk = "��"
							   else
							        old_chk = ""
						    end if
							if rs("f_holt") = "Y" then
							        holt_chk = "��"
							   else
							        holt_chk = ""
						    end if
							if rs("f_woman") = "Y" then
							        woman_chk = "��"
							   else
							        woman_chk = ""
						    end if
							if rs("f_single") = "Y" then
							        single_chk = "��"
							   else
							        single_chk = ""
						    end if
							if rs("f_disab") = "Y" then
							        disab_chk = "��"
							   else
							        disab_chk = ""
						    end if
							if rs("f_children") = "Y" then
							        children_chk = "��"
							   else
							        children_chk = ""
						    end if

							hap_nts_donation = rs("d_poli_nts") + rs("d_poli10_nts") + rs("d_law_nts") + rs("d_ji_nts")
		                    hap_oth_donation = rs("d_poli_other") + rs("d_poli10_other") + rs("d_law_other") + rs("d_ji_other")

							if bon_person_no = rs("f_person_no") and rs_year("y_emp_no") = rs("f_emp_no") then
							       hap_nts_nhis = rs_year("y_nhis_amt") + rs_year("y_epi_amt") + rs_year("y_longcare_amt")
								   hap_oth_insuran = rs("i_ilban_other") + rs("i_disab_other")
								   hap_nts_insuran = rs("i_ilban_nts") + rs("i_disab_nts")
								   hap_oth_nhis = 0
							   else
								   hap_nts_nhis = 0
								   hap_oth_nhis = 0
								   hap_nts_insuran = rs("i_ilban_nts") + rs("i_disab_nts")
		                           hap_oth_insuran = rs("i_ilban_other") + rs("i_disab_other")
							end if
	           			%>
							<tr>
                                <td style=" border-top:2px solid #515254;"><%=rel_chk%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=rs("f_family_name")%>&nbsp;</td>
                                <td colspan="2" style=" border-top:2px solid #515254;"><%=basic_chk%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=old_chk%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=holt_chk%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;">����û�ڷ�</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(hap_nts_nhis,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(hap_nts_insuran,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("m_nts_amt"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("e_nts_amt"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("c_credit_nts"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("c_direct_nts"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("c_cash_nts"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("c_market_nts"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(rs("c_transit_nts"),0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(hap_nts_donation,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td><%=rs("f_national")%>&nbsp;</td>
                                <td><%=f_person_no%>&nbsp;</td>
                                <td><%=woman_chk%>&nbsp;</td>
                                <td><%=single_chk%>&nbsp;</td>
                                <td><%=disab_chk%>&nbsp;</td>
                                <td><%=children_chk%>&nbsp;</td>
                                <td>�׹����ڷ�</td>
                                <td class="right"><%=formatnumber(hap_oth_nhis,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_oth_insuran,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("m_other_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("e_other_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("c_credit_other"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("c_direct_other"),0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("c_market_other"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("c_transit_other"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_oth_donation,0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
                       </table>
                       </td>
                      </tr>
                </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="100%" valign="top">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                       <colgroup>
							   <col width="4%" >
                               <col width="*" >
							   <col width="10%" >
							   <col width="8%" >
                               <col width="8%" >

							   <col width="14%" >
                               <col width="14%" >
                               <col width="14%" >
                               <col width="14%" >
                        </colgroup>
                        <thead>
                            <tr>
								<td colspan="9" scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
							</tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:2px solid #515254;">����</th>
                              <th colspan="3" style=" border-bottom:2px solid #515254;">�����</th>
                              <th style=" border-bottom:2px solid #515254;">���ⱸ��</th>
                              <th style=" border-bottom:2px solid #515254;">�ݾ�</th>
                              <th style=" border-bottom:2px solid #515254;">�ѵ���</th>
                              <th style=" border-bottom:2px solid #515254;">������</th>
						    </tr>
                            <tr>
							  <th rowspan="5">��.���ݺ�������</th>
                              <th rowspan="5">���ݺ����<br>(���ο���,����������,���ο���,����������,���ݰ��µ�)</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3;">���ο��ݺ����</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(b_nps,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(b_nps_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right" ><%=formatnumber(y_nps_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(y_nps_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ο��ݺ���� ���� �������ݺ����</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(y_nps_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(y_nps_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(y_nps_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(y_nps_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3;">���ݺ���� ��</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(a_tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(a_tot_tax,0)%>&nbsp;</td>
						    </tr>


                            <tr>
							  <th rowspan="13" style=" border-top:2px solid #515254;">��. Ư���ҵ����</th>
                              <th rowspan="5" style=" border-top:2px solid #515254;">�����</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">���ΰǰ�����<br>(��������纸�� ����)</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">�����</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(b_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">����</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(b_nhis_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right" ><%=formatnumber(y_nhis_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(y_nhis_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��뺸��</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(b_epi,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(b_epi_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��(��)�ٹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(y_epi_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(y_epi_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3;">����� ��</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(i_tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(i_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="8" style=" border-top:2px solid #515254; border-left:1px solid #e3e3e3;">�����ڱ�</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">�����������Ա�</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">����������</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">�����ݻ�ȯ��</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(h_lender_amt,0)%>&nbsp;</td>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">�ۼ���� ����</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(h_lender_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">������ ����</th>
                              <td class="right"><%=formatnumber(h_person_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_person_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="5" style=" border-left:1px solid #e3e3e3;">��������������Ա�</th>
                              <th rowspan="3" style=" border-bottom:1px solid #e3e3e3;">2011�� ����<br>���Ժ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">15��̸�</th>
                              <th rowspan="5" >���ڻ�ȯ��</th>
                              <td class="right"><%=formatnumber(h_long15_amt,0)%>&nbsp;</td>
                              <th rowspan="5" style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(h_long15_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15�� ~ 29��</th>
                              <td class="right"><%=formatnumber(h_long29_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_long29_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">30��</th>
                              <td class="right"><%=formatnumber(h_long30_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_long30_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; ">2012�� ����<br>���Ժ�<br>(15���̻�)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����ݸ�.���ġ��ȯ����</th>
                              <td class="right"><%=formatnumber(h_fixed_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_fixed_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; ">��Ÿ����</th>
                              <td class="right"><%=formatnumber(h_other_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">�����ڱ� ������ ��</th>
                              <th style=" border-top:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(h_tot_amt,0)%>&nbsp;</td>
                              <th >&nbsp;</th>
                              <td class="right"><%=formatnumber(h_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="21" style=" border-top:2px solid #515254;">��. �� �� �� �� �� �� ��</th>
                              <th colspan="4" style=" border-top:2px solid #515254; border-bottom:1px solid #e3e3e3;">���ο�������(2000�� 12�� 31�� ���� ����)</th>
                              <th style=" border-top:2px solid #515254; border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(tot_2000,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">���Ծ�40%��(72����)</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(tax_2000,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ұ�����һ���� �����α�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >���Աݾ�</th>
                              <td class="right" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ø�������</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3; ">û������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���Աݾ�</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_cheng,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�ٷ������ø�������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���Աݾ�</th>
                              <td class="right" ><%=formatnumber(tot_gunro,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_gunro,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">����û����������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���Աݾ�</th>
                              <td class="right" ><%=formatnumber(tot_jutak,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_jutak,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">���ø������� �ҵ���� ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������� ���ڵ�</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3; ">2012�� ���ڡ����ں�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���ڡ����ڱݾ�</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_cheng,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">2013�� ���ڡ����ں�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���ڡ����ڱݾ�</th>
                              <td class="right" ><%=formatnumber(tot_gunro,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_gunro,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">2014�� ���� ���ڡ����ں�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���ڡ����ڱݾ�</th>
                              <td class="right" ><%=formatnumber(tot_jutak,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(tax_jutak,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�������� ���� �� �ҵ���� ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>



                            <tr>
							  <th rowspan="10" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ſ�ī�� �� ����</th>
                              <th colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3; ">��ſ�ī��(������塤���߱����������)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�����ҡ�����ī��(������塤���߱����������)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left"style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�����ݿ�����(������塤���߱����������)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">������������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">����߱������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">��(��+��+��+��+��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�캻�� �ſ�ī��� ����(2013��)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">���� �ſ�ī��� ����(2014��)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">��� �߰������� ����(2013��)<br>-����,����,���ݿ�����,�������,���߱���</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�ﺻ�� �߰������� ����(2014���Ϲݱ�)<br>-����,����,���ݿ�����,�������,���߱���</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>


                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�츮�������� �⿬��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�⿬�ݾ�</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�츮�������� ��α�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">��αݾ�</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">��������߼ұ�� �ٷ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ӱݻ谨��</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�� �� ��� ���� ���ڻ�ȯ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���ڻ�ȯ��</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">�������������������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">���Աݾ�</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">�ۼ���� ����</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                        </thead>
                        <tbody>
						</tbody>
                       </table>
                       </td>
                      </tr>
                </table>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="100%" valign="top">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                       <colgroup>
							   <col width="4%" >
                               <col width="4%" >
                               <col width="4%" >
                               <col width="4%" >
							   <col width="8%" >
                               <col width="8%" >

							   <col width="14%" >
                               <col width="12%" >
                               <col width="12%" >
                               <col width="12%" >
                               <col width="6%" >
                               <col width="12%" >
                        </colgroup>
                        <thead>
                            <tr>
								<td colspan="12" scope="col" style=" border-bottom:2px solid #515254;">&nbsp;</td>
							</tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:2px solid #515254;">����</th>
                              <th colspan="4" style=" border-bottom:2px solid #515254;">���װ��顤������</th>
                              <th colspan="6" style=" border-bottom:2px solid #515254;">���װ��顤������</th>
						    </tr>
                            <tr>
							  <th rowspan="35" >��. ���װ���� ����</th>
                              <th rowspan="5" style="border-bottom:1px solid #e3e3e3;">���װ���</th>
                              <th rowspan="4" style="border-bottom:1px solid #e3e3e3;">�ܱ��αٷ���</th>
                              <th colspan="2" style="border-bottom:1px solid #e3e3e3;">�Ա�����</th>
                              <td colspan="7">[ ]���ΰ� ���� [ ]������԰�� [ ]������Ư�����ѹ����� ���� [ ]�������� �� ����</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������԰�� �Ǵ� �ٷ�������</th>
                              <td ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; " >����Ⱓ ������</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ܱ��� �ٷμҵ濡 ���� ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td colspan="2" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ٷμҵ濡 ���� �������� �� ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td colspan="2" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�߼ұ�� ����� ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�����</th>
                              <td colspan="2" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >����Ⱓ ������</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="30" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���װ���</th>
                              <th colspan="4" style="border-bottom:1px solid #e3e3e3;">�� �� �� ��</th>
                              <th colspan="2" style="border-bottom:1px solid #e3e3e3;">��</th>
                              <th style="border-bottom:1px solid #e3e3e3;">�ѵ���</th>
                              <th style="border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <th style="border-bottom:1px solid #e3e3e3;">������</th>
                              <th style="border-bottom:1px solid #e3e3e3;">��������</th>
						    </tr>
                            <tr>
                              <th rowspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ݰ���</th>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���б���ΰ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >���Աݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th rowspan="3" style=" border-bottom:1px solid #e3e3e3; " >�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="4">12%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ٷ��������޿� ��������� ���� ��������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >���Աݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >���Աݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ݰ��� ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>

                            <tr>
                              <th rowspan="17" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">Ư�����װ���</th>
                              <th rowspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����</th>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���强</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�����</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >100����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="3">12%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��������뺸�强</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�����</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >100����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����� ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�Ƿ��</th>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���Ρ�65���̻��ڡ������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�����</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="3">15%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�� ���� ���������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�����</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�Ƿ�� ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="6" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������</th>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ҵ��� ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������(���п�����)</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="6">15%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������ �Ƶ�(<%=adong_cnt%> ��)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >��ġ�����п����</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >1��� 300����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ʡ��ߡ�����б�(<%=adong_cnt%> ��)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >1��� 300����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���л�(���п� ������)(<%=adong_cnt%> ��)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >1��� 300����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����(<%=adong_cnt%> ��)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >Ư��������</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������ ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��α�</th>
                              <th rowspan="2" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��ġ�ڱݱ�α�</th>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">10��������</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >��αݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th rowspan="4" style=" border-bottom:1px solid #e3e3e3; " >�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >100/110</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">10�����ʰ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >��αݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >15%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������α�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >��αݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >25%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������α�</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >��αݾ�</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >25%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��α� ��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th >&nbsp;</th>
                              <td class="right" style=" border-bottom:1px solid #e3e3e3; "><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</td>
                              <td class="right" style=" border-bottom:1px solid #e3e3e3; "><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>

                            <tr>
                              <th rowspan="6" colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ܱ����μ���</th>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ܿ�õ�ҵ�</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >������(��ȭ)</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >������(��ȭ)</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
                              <th colspan="2" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >��������</th>
                              <td colspan="2" class="right"><%=de_tax_nation%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >������</th>
                              <td colspan="2" class="right"><%=de_tax_date%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >��û��������</th>
                              <td colspan="2" class="right"><%=de_tax_nation%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >���ܱٹ�ó</th>
                              <td colspan="2" class="right"><%=de_tax_date%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >�ٹ��Ⱓ</th>
                              <td colspan="2" class="right"><%=de_tax_nation%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >��å</th>
                              <td colspan="2" class="right"><%=de_tax_date%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����ڱ����Ա����ڼ��װ���</th>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >���ڻ�ȯ��</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >30%</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������װ���</th>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >������</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >10%</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
								<td colspan="12" scope="col" class="left" style="border-top:2px solid #515254;">�Ű����� ���ҵ漼���� ��140���� ���� ���� ������ �Ű��ϸ�,<br>�� ������ ����� �����Ͽ��� �Ű����� �˰� �ִ� ��� �׷��θ� ��Ȯ�ϰ� �������� Ȯ���մϴ�.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 �� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��û�� : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(���� �Ǵ� ��)<br></td></td>
							</tr>
                        </thead>
                        <tbody>
						</tbody>
                       </table>
                       </td>
                      </tr>
                </table>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="100%" valign="top">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                       <colgroup>
							   <col width="13%" >
                               <col width="14%" >
                               <col width="20%" >
                               <col width="13%" >
							   <col width="20%" >
                               <col width="20%" >
                        </colgroup>
                        <thead>
                            <tr>
								<td colspan="6" scope="col" class="left">��. �߰� ���� ����</td>
							</tr>
                            <tr>
								<td colspan="5" scope="col" class="left" >1. �ܱ��αٷ��� ���ϼ��������û�� ���� ����(�� �Ǵ� ���� �����ϴ�)</td>
                                <td scope="col" >���� (&nbsp;&nbsp;&nbsp;)</td>
							</tr>
                            <tr>
								<th rowspan="2" scope="col" style=" border-top:1px solid #e3e3e3; "  >2. ��(��)�ٹ��� ��</th>
                                <th scope="col" style="border-top:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >��(��)�ٹ�����</th>
                                <td scope="col" ><%=de_tax_date%>&nbsp;</td>
                                <th scope="col" style="border-top:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >��(��)�޿��Ѿ�</th>
                                <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                                <td rowspan="2" scope="col" >��(��)�ٹ��� �ٷμҵ�<br>��õ¡�������� ����(&nbsp;&nbsp;)</td>
							</tr>
                            <tr>
                                <th scope="col" style=" border-left:1px solid #e3e3e3; " >����ڵ�Ϲ�ȣ</th>
                                <td scope="col" ><%=de_tax_date%>&nbsp;</td>
                                <th scope="col" >��(��)��������</th>
                                <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
							</tr>
                            <tr>
								<td colspan="3" scope="col" class="left">3. ���ݡ����� �� �ҵ桤���� �������� ���⿩��(�� �Ǵ� ���� �����ϴ�)</td>
                                <td colspan="3" scope="col" class="left">����(&nbsp;&nbsp;) �� ���ݰ���, ���ø������� �� �ҵ桤���װ����� ��û�� ��� �ش� ������ �����ؾ� �մϴ�.</td>
							</tr>
                            <tr>
								<td colspan="3" scope="col" class="left">4. �����ס������� �� �����������ӱ� �����ݻ�ȯ�� �ҵ���� ���� ���⿩��(�� �Ǵ� ���� �����ϴ�)</td>
                                <td colspan="3" scope="col" class="left">����(&nbsp;&nbsp;) �� �����ס������� �� �����������ӱ� �����ݻ�ȯ�� �ҵ������ ��û�� ��� �ش� ������ �����ؾ� �մϴ�.</td>
							</tr>
                            <tr>
								<td colspan="2" scope="col" class="left">5. �� ���� �߰� ���� ����</td>
                                <td colspan="4" scope="col" >���Ƿ�����޸���(&nbsp;&nbsp;), ���αݸ���(&nbsp;&nbsp;), ��ҵ���� ������(&nbsp;&nbsp;)</td>
							</tr>
                        </thead>
                        <tbody>
						</tbody>
                       </table>
                       </td>
                      </tr>
                </table>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="100%" valign="top">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                       <colgroup>
							   <col width="100%" >
                        </colgroup>
                        <thead>
                            <tr>
								<td scope="col" style=" border-bottom:2px solid #515254;">���ǻ���</td>
							</tr>
                            <tr>
								<td scope="col" class="left" >1. �ٷμҵ��ڰ� ��(��)�ٹ��� �ٷμҵ��� ��õ¡���ǹ��ڿ��� �Ű����� ���� ��쿡�� �ٷμҵ��� ������ ���ռҵ漼 �Ű� �ؾ� �ϸ�, �Ű����� ���� ��� ���꼼 �ΰ� �� �������� �����ϴ�.<br><br>2. �� �ٹ����� ���ݺ���ᡤ���ΰǰ������ �� ��뺸��� ���� �Ű����� �ۼ����� �ʾƵ� �˴ϴ�.<br><br>3. �����ݾ׶��� �ٷ��ڰ� ��õ¡���ǹ��ڿ��� �����ϴ� ��� ���� ���� �� �ֽ��ϴ�.</td>
                            </tr>
                        </thead>
                        <tbody>
						</tbody>
                       </table>
                       </td>
                      </tr>
                </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_tax_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_tax_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">�ҵ�����Ű� ���</a>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>

