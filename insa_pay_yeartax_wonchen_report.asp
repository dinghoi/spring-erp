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

w = datepart("w",curr_date)
response.write(w)

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
emp_addr = cstr(rs_emp("emp_sido")) + " " + cstr(rs_emp("emp_gugun")) + " " + cstr(rs_emp("emp_dong")) + " " + cstr(rs_emp("emp_addr"))	 
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
	   y_from_date = rs_year("y_from_date")
	   y_to_date = rs_year("y_to_date")
	   
	   y_total_pay = rs_year("y_total_pay")
	   y_total_bonus = rs_year("y_total_bonus")
	   y_other_pay = rs_year("y_other_pay")
   else
       y_nps_amt = 0
	   y_nhis_amt = 0
	   y_longcare_amt = 0
	   y_epi_amt = 0
	   y_householder = "N"
	   y_national = ""
	   y_live = "Y"
	   y_change = "N"
end if

y_nps_tax = y_nps_amt
y_nhis_amt = y_nhis_amt + y_longcare_amt
y_nhis_tax = y_nhis_amt
y_epi_tax = y_epi_amt

if rs_year("y_woman") = "Y" then
        y_woman = "��"
   else
        y_woman = ""
end if
if rs_year("y_single") = "Y" then
        y_single = "��"
   else
        y_single = ""
end if

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

sum_nts_nhis =  rs_year("y_nhis_amt") + rs_year("y_epi_amt") + rs_year("y_longcare_amt")

b_nps = 0
b_nhis = 0
b_longcare = 0
b_epi = 0
b_pay = 0
b_bonus = 0
b_deem_bonus = 0
Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
'Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
	   tot_pay = tot_pay + rs_bef("b_pay") + rs_bef("b_bonus") + rs_bef("b_deem_bonus")
	   b_nps = b_nps + rs_bef("b_nps")
	   b_nhis = b_nhis + rs_bef("b_nhis")
	   b_longcare = b_longcare + rs_bef("b_longcare")
	   b_epi = b_epi + rs_bef("b_epi")
	   b_company =  rs_bef("b_company")
	   b_company_no =  rs_bef("b_company_no")
	   b_from_date =  rs_bef("b_from_date")
	   b_to_date =  rs_bef("b_to_date")
	   
	   b_pay =  rs_bef("b_pay")
	   b_bonus =  rs_bef("b_bonus")
	   b_deem_bonus =  rs_bef("b_deem_bonus")
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
	  owner_person_no = "�����"
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

title_line = "�ٷμҵ� ��õ¡�� ������"
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
				<form action="insa_pay_yeartax_wonchen_report.asp" method="post" name="frm">
				<div class="gView">
				<table border="0" cellpadding="0" cellspacing="0" class="tableList">
				  <tr>
				    <td width="60%" style="font-size:18px;"><strong>[&nbsp;&nbsp;]�ٷμҵ� ��õ¡��������<br>[&nbsp;&nbsp;]�ٷμҵ� ���� ����(��)</strong></td>
				    <td width="*"><table cellspacing="0" cellpadding="0" class="tableList">
				      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">���ֱ���</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">������1/�������2</td>
			          </tr>
                      <tr>
				        <td class="center" width="20%" style=" border-left:1px solid #000000;">��������</td>
                        <td class="center" width="20%">&nbsp;</td>
				        <td class="center" width="30%">���������ڵ�</td>
                        <td class="center" width="30%" style=" border-right:1px solid #000000;">&nbsp;</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">�����ܱ���</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">������1/�ܱ���9</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">�ܱ��δ��ϼ�������</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">��1/��2</td>
			          </tr>
                      <tr>
				        <td class="center" width="20%" style=" border-left:1px solid #000000;">����</td>
				        <td class="center" width="30%">&nbsp;</td>
                        <td class="center" width="20%">�����ڵ�</td>
                        <td class="center" width="30%" style=" border-right:1px solid #000000;">&nbsp;</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">�����ֿ���</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">������1 �����2,</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">�������� ����</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">��ӱٷ�1, �ߵ����2</td>
			          </tr>
				      </table>
                    </td>
			      </tr>
				  </table>


                  <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
							<col width="20%" >
						</colgroup>
						<thead>
                            <tr>
							  <th rowspan="3" >¡���ǹ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">����θ�(��ȣ)</th>
                              <td ><%=company_name%></td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���ǥ��(����)</th>
                              <td ><%=owner_name%></td>
						    </tr>
                            <tr>
							  <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����ڵ�Ϲ�ȣ</th>
                              <td ><%=trade_no%></td>
                              <th >���ֹε�Ϲ�ȣ</th>
                              <td ><%=owner_person_no%></td>
						    </tr>
                            <tr>
							  <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������(�ּ�)</th>
                              <td class="left" colspan="3"><%=addr_name%></td>
						    </tr>
                            <tr>
							  <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">�ҵ���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�켺��</th>
                              <td ><%=emp_name%></td>
                              <th style=" border-top:1px solid #e3e3e3;">���ֹε�Ϲ�ȣ</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th style=" border-left:1px solid #e3e3e3;">���ּ�</th>
                              <td class="left" colspan="3"><%=emp_addr%></td>
						    </tr>
						</thead>
					</table>

                  <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
							<col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th rowspan="12" >��<br>��<br>��<br>ó<br>��<br>��<br>��<br>��<br>��</th>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">����</th>
                              <th >��(��)</th>
                              <th >��(��)</th>
                              <th >��(��)</th>
                              <th >��������</th>
                              <th >�հ�</th>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ٹ�ó��</th>
                              <td ><%=company_name%></td>
                              <td ><%=b_company%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����ڵ�Ϲ�ȣ</th>
                              <td ><%=trade_no%></td>
                              <td ><%=b_company_no%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ٹ��Ⱓ</th>
                              <td ><%=y_from_date%>~<%=y_to_date%></td>
                              <td ><%=b_from_date%>~<%=b_to_date%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����Ⱓ</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�޿�</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��</th>
                              <td class="right"><%=formatnumber(y_total_bonus,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_bonus,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������</th>
                              <td class="right"><%=formatnumber(y_other_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_deem_bonus,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ֽĸż����ñ� �������</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�츮�������������</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ӿ������ҵ�ݾ� �ѵ��ʰ���</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
						</thead>
					</table>

                  <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
                            <col width="4%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
							<col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th rowspan="11" >��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��</th>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ܱٷ�</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">M0X</th>
                              <td ><%=company_name%></td>
                              <td ><%=b_company%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�߰��ٷμ���</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">O0X</th>
                              <td ><%=trade_no%></td>
                              <td ><%=b_company_no%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���.��������</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">Q0X</th>
                              <td ><%=y_from_date%>~<%=y_to_date%></td>
                              <td ><%=b_from_date%>~<%=b_to_date%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����������</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">H0X</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">-5</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">-6</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(y_total_bonus,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_bonus,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">~</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(y_other_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_deem_bonus,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">-25</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ú�������</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">Y22</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������ҵ��</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����ҵ��</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
  					    </thead>
					</table>
                  <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
                            <col width="10%" >
                            <col width="8%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
							<col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th rowspan="8" >��<br>��<br>��<br>��<br>��</th>
                              <th colspan="4" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">����</th>
                              <th >�ҵ漼</th>
                              <th >����ҵ漼</th>
                              <th >�����Ư����</th>
						    </tr>
                            <tr>
							  <th colspan="4" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��������</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="4" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ⳳ�μ���</th>
                              <th rowspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��(��)�ٹ���<br>(�������׶��� ���ױ���)</th>
                              <th rowspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����ڵ�Ϲ�ȣ</th>
                              <td style=" border-left:1px solid #e3e3e3;"><%=b_company_no%></td>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <td style=" border-left:1px solid #e3e3e3;"><%=b_company_no%></td>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
						    <tr>
                              <td style=" border-left:1px solid #e3e3e3;"><%=b_company_no%></td>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��(��)�ٹ���</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            
                            <tr>
							  <th colspan="4" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����Ư�ʼ���</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th colspan="4" class="left" style=" border-left:1px solid #e3e3e3;">����¡������</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                                <td colspan="8" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">���� ��õ¡����(�ٷμҵ�)�� ���� ����(����)�մϴ�.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 �� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��û�� : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(���� �Ǵ� ��)<br><strong>��������</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����.</td>
                            </tr>
						</thead>
					</table>
                    
                  <table width="100%" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="14%" >
                            
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th rowspan="40" >��<br>��<br>��<br>��<br>��</th>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;" >�ѱ޿�(�ٸ�, �ܱ��δ��ϼ������� �ÿ��� ���� �ٷμҵ�)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">�ҵ���� �����ѵ� �ʰ���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="6" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ٷμҵ����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">���ռҵ� ����ǥ��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="6" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ٷμҵ�ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">���⼼��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="23" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��<br>��<br>��<br>��<br>��<br>��</th>
                              <th rowspan="3" style=" border-bottom:1px solid #e3e3e3;">��<br>��<br>��<br>��</th>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="5" style=" border-bottom:1px solid #e3e3e3;" >��<br>��<br>��<br>��</th>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">�ҵ漼��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" >����Ư�����ѹ�(��30�� ����)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ξ簡��(<%=rs_year("y_support_cnt")%> ��)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">����Ư�����ѹ� ��30��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��<br>��<br>��<br>��</th>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">��ο��(<%=rs_year("y_support_cnt")%> ��)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">��������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����(<%=rs_year("y_support_cnt")%> ��)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" >���װ��� ��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�γ���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="31" style=" border-bottom:1px solid #e3e3e3;">��<br>��<br>��<br>��</th>
                              <th class="left" colspan="5" style=" border-top:1px solid #e3e3e3;">�ٷμҵ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�Ѻθ���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">�ڳ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="5" style=" border-left:1px solid #e3e3e3;">��<br>��<br>��<br>��<br>��<br>��<br>��</th>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">���ο��ݺ����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="6" style=" border-bottom:1px solid #e3e3e3;">��<br>��<br>��<br>��</th>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">���б���ΰ���</th>
                              <th >�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������ݺ�������</th>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">����������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ο���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">�ٷ��������޿�������� ���� ��������</th>
                              <th >�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�縳�б�����������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������ü������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">��������</th>
                              <th >�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="11" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">Ư<br>��<br>��<br>��<br>��<br>��</th>
                              <th class="left" rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����</th>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">�ǰ������<br>(��������纸�������)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��뺸���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="16" style=" border-bottom:1px solid #e3e3e3;">Ư<br>��<br>��<br>��<br>��<br>��</th>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">���强�����</th>
                              <th >�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="7" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����ڱ�</th>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����������Աݿ����ݻ�ȯ��</th>
                              <th class="left" colspan="2" style=" border-bottom:1px solid #e3e3e3;">������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="2" style=" border-bottom:1px solid #e3e3e3;">������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">�Ƿ��</th>
                              <th >�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="5" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��������������Ա����ڻ�ȯ��</th>
                              <th class="left" rowspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2011�� ���� ���Ժ�</th>
                              <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15�� �̸�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15��~29��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">30���̻�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2012�� ���� ���Ժ�<br>(15���̻�)</th>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">��ݸ�.���ġ��ȯ����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="8" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��<br>��<br>��</th>
                              <th class="left" rowspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��ġ�ڱݱ�α�</th>
                              <th class="left" rowspan="2" style=" border-bottom:1px solid #e3e3e3;">10��������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">�� ���� ����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��α�(�̿���)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" style=" border-bottom:1px solid #e3e3e3;">10�����ʰ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="6" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�����ҵ�ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="2" style=" border-bottom:1px solid #e3e3e3;">������α�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="13" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��<br>��<br>��<br>��<br>��<br>��<br>��</th>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">���ο�������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">�ұ��.�һ���� �����α�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="2" style=" border-bottom:1px solid #e3e3e3;">������α�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="3" colspan="2" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ø��� ���� �ҵ����</th>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">û������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">����û����������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">�ٷ������ø�������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">ǥ�ؼ��װ���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������������ ��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">�������հ���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ſ�ī�� �� ����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">�������Ա�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�츮�������� �⿬��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">�ܱ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�츮�������� ��α�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="4" style=" border-bottom:1px solid #e3e3e3;">������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ݾ�</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������� �߼ұ�� �ٷ���</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���װ�����</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�񵷾ȵ�� ���� ���ڻ�ȯ��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">���� ���� ��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������������������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th colspan="5" style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right">&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�� ���� �ҵ���� ��</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">��������</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
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
                                <td class="left" colspan="18" scope="col" style=" border-bottom:1px solid #e3e3e3;">�ҵ桤���װ��� ��(���������׸��� �ش���� ��ǥ��(����� �ش�ô� �ش��ڵ� ����)�� �ϸ�, ���� �ҵ���������װ��� �׸��� ������ ���Ͽ� ���� ������ �ݾ��� �����ϴ�)</td>
							</tr>
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
                            
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_wonchen_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_tax_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">��õ¡�������� ���</a>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

