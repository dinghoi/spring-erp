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
        y_woman = "○"
   else
        y_woman = ""
end if
if rs_year("y_single") = "Y" then
        y_single = "○"
   else
        y_single = ""
end if

y_national_code = "001"

if y_householder = "Y" then
       householder = "[∨]세대주 [ ]세대원"
   else
       householder = "[ ]세대주 [∨]세대원"
end if

if y_live = "Y" then
       yy_live = "[∨]거주자 [ ]비거주자"
   else
       yy_live = "[ ]거주자 [∨]비거주자"
end if

if y_change = "N" then
       yy_change = "[∨]전년과동일 [ ]변동"
   else
       yy_change = "[ ]전년과동일 [∨]변동"
end if

'신용카드
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

'연금보험
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

'보험료
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

'주택자금
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

'연금저축 
s_id = "연금저축"
tot_2000 = 0
tot_2001 = 0
tot_endi = 0
Sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
rs_savi.Open Sql, Dbconn, 1
Set rs_savi = DbConn.Execute(SQL)
do until rs_savi.eof
       if rs_savi("s_type") = "개인연금저축(2000년이전)" then 
	           tot_2000 = tot_2000 + rs_savi("s_amt")
		  elseif rs_savi("s_type") = "연금저축(2001년이후)" then 
	                  tot_2001 = tot_2001 + rs_savi("s_amt")
			     elseif rs_savi("s_type") = "퇴직연금소득공제" then 
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

'기타공제/투자조합출자
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

 s_id = "주택마련저축" 
      tot_cheng = 0
      tot_jutak = 0
      tot_gunro = 0
	  tot_jangi = 0
      Sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
      rs_savi.Open Sql, Dbconn, 1
      Set rs_savi = DbConn.Execute(SQL)
      do until rs_savi.eof
            if rs_savi("s_type") = "청약저축" then 
	                 tot_cheng = tot_cheng + rs_savi("s_amt")
		       elseif rs_savi("s_type") = "주택청약종합저축" then 
	                        tot_jutak = tot_jutak + rs_savi("s_amt")
			          elseif rs_savi("s_type") = "근로자주택마련저축" then 
	                              tot_gunro = tot_gunro + rs_savi("s_amt")
							 elseif rs_savi("s_type") = "장기주택마련저축" then 
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
if emp_company = "케이원정보통신" then
      company_name = "(주)" + "케이원정보통신"
	  owner_name = "김승일"
	  owner_person_no = "김승일"
	  addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif emp_company = "휴디스" then
              company_name = "(주)" + "휴디스"
			  owner_name = "김한종"
	          addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif emp_company = "케이네트웍스" then
                     company_name = "케이네트웍스" + "(주)"
					 owner_name = "이중원"
	                 addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif emp_company = "에스유에이치" then
                        company_name = "(주)" + "에스유에이치"	
						owner_name = "박미애"
	                    addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

c_hap1 = 0

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = "근로소득 원천징수 영수증"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
					alert ("귀속년도를 입력하세요.");
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
				    <td width="60%" style="font-size:18px;"><strong>[&nbsp;&nbsp;]근로소득 원천징수영수증<br>[&nbsp;&nbsp;]근로소득 지급 명세서(안)</strong></td>
				    <td width="*"><table cellspacing="0" cellpadding="0" class="tableList">
				      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">거주구분</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">거주자1/비거주자2</td>
			          </tr>
                      <tr>
				        <td class="center" width="20%" style=" border-left:1px solid #000000;">거주지국</td>
                        <td class="center" width="20%">&nbsp;</td>
				        <td class="center" width="30%">거주지국코드</td>
                        <td class="center" width="30%" style=" border-right:1px solid #000000;">&nbsp;</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">내·외국인</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">내국인1/외국인9</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">외국인단일세율적용</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">여1/부2</td>
			          </tr>
                      <tr>
				        <td class="center" width="20%" style=" border-left:1px solid #000000;">국적</td>
				        <td class="center" width="30%">&nbsp;</td>
                        <td class="center" width="20%">국적코드</td>
                        <td class="center" width="30%" style=" border-right:1px solid #000000;">&nbsp;</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">세대주여부</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">세대주1 세대원2,</td>
			          </tr>
                      <tr>
				        <td class="center" colspan="2" style=" border-left:1px solid #000000;">연말정산 구분</td>
				        <td class="center" colspan="2" style=" border-right:1px solid #000000;">계속근로1, 중도퇴사2</td>
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
							  <th rowspan="3" >징수의무자</th>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">①법인명(상호)</th>
                              <td ><%=company_name%></td>
                              <th style=" border-bottom:1px solid #e3e3e3;">②대표자(성명)</th>
                              <td ><%=owner_name%></td>
						    </tr>
                            <tr>
							  <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">③사업자등록번호</th>
                              <td ><%=trade_no%></td>
                              <th >④주민등록번호</th>
                              <td ><%=owner_person_no%></td>
						    </tr>
                            <tr>
							  <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">⑤소재지(주소)</th>
                              <td class="left" colspan="3"><%=addr_name%></td>
						    </tr>
                            <tr>
							  <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">소득자</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">⑥성명</th>
                              <td ><%=emp_name%></td>
                              <th style=" border-top:1px solid #e3e3e3;">⑦주민등록번호</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th style=" border-left:1px solid #e3e3e3;">⑧주소</th>
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
							  <th rowspan="12" >Ⅰ<br>근<br>무<br>처<br>별<br>소<br>득<br>명<br>세</th>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">구분</th>
                              <th >주(현)</th>
                              <th >종(전)</th>
                              <th >종(전)</th>
                              <th >납세조합</th>
                              <th >합계</th>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">근무처명</th>
                              <td ><%=company_name%></td>
                              <td ><%=b_company%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">사업자등록번호</th>
                              <td ><%=trade_no%></td>
                              <td ><%=b_company_no%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">근무기간</th>
                              <td ><%=y_from_date%>~<%=y_to_date%></td>
                              <td ><%=b_from_date%>~<%=b_to_date%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">감면기간</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">급여</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">상여</th>
                              <td class="right"><%=formatnumber(y_total_bonus,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_bonus,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">인정상여</th>
                              <td class="right"><%=formatnumber(y_other_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_deem_bonus,0)%>&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주식매수선택권 행사이익</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">우리사주조합인출금</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">임원퇴직소득금액 한도초과액</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">계</th>
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
							  <th rowspan="11" >Ⅱ<br>비<br>과<br>세<br>및<br>감<br>면<br>소<br>득<br>명<br>세</th>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">국외근로</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">M0X</th>
                              <td ><%=company_name%></td>
                              <td ><%=b_company%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">야간근로수당</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">O0X</th>
                              <td ><%=trade_no%></td>
                              <td ><%=b_company_no%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">출산.보육수당</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">Q0X</th>
                              <td ><%=y_from_date%>~<%=y_to_date%></td>
                              <td ><%=b_from_date%>~<%=b_to_date%></td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">연구보조비</th>
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
							  <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">수련보조수당</th>
                              <th class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">Y22</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">비과세소득계</th>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">감면소득계</th>
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
							  <th rowspan="8" >Ⅲ<br>세<br>액<br>명<br>세</th>
                              <th colspan="4" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">구분</th>
                              <th >소득세</th>
                              <th >지방소득세</th>
                              <th >농어촌특별세</th>
						    </tr>
                            <tr>
							  <th colspan="4" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">결정세액</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="4" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기납부세액</th>
                              <th rowspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">종(전)근무지<br>(결정세액란의 세액기재)</th>
                              <th rowspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">사업자등록번호</th>
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
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주(현)근무지</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            
                            <tr>
							  <th colspan="4" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">납부특례세액</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th colspan="4" class="left" style=" border-left:1px solid #e3e3e3;">차감징수세액</th>
                              <td class="right"><%=formatnumber(y_total_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(b_pay,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                                <td colspan="8" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">위의 원천징수액(근로소득)을 정히 영수(지급)합니다.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 년 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;월 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;일<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;신청인 : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(서명 또는 인)<br><strong>세무서장</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;귀하.</td>
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
							  <th rowspan="40" >Ⅳ<br>정<br>산<br>명<br>세</th>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;" >총급여(다만, 외국인단일세율적용 시에는 연간 근로소득)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">소득공제 종합한도 초과액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="6" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">근로소득공제</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">종합소득 과세표준</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="6" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">근로소득금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">산출세액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="23" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">종<br>합<br>소<br>득<br>공<br>제</th>
                              <th rowspan="3" style=" border-bottom:1px solid #e3e3e3;">기<br>본<br>공<br>제</th>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">본인</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="5" style=" border-bottom:1px solid #e3e3e3;" >세<br>액<br>감<br>면</th>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">소득세법</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">배우자</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" >조세특례제한법(제30조 제외)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">부양가족(<%=rs_year("y_support_cnt")%> 명)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">조세특례제한법 제30조</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">추<br>가<br>공<br>제</th>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">경로우대(<%=rs_year("y_support_cnt")%> 명)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">조세조약</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장애인(<%=rs_year("y_support_cnt")%> 명)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" >세액감면 계</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">부녀자</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="31" style=" border-bottom:1px solid #e3e3e3;">세<br>액<br>공<br>제</th>
                              <th class="left" colspan="5" style=" border-top:1px solid #e3e3e3;">근로소득</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">한부모가족</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">자녀</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="5" style=" border-left:1px solid #e3e3e3;">연<br>금<br>보<br>험<br>료<br>공<br>제</th>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">국민연금보험료</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="6" style=" border-bottom:1px solid #e3e3e3;">연<br>금<br>계<br>좌</th>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">과학기술인공제</th>
                              <th >공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">공적연금보험료공제</th>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">공무원연금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">군인연금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">근로자퇴직급여보장법에 따른 퇴직연금</th>
                              <th >공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">사립학교교직원연금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">별정우체국연금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">연금저축</th>
                              <th >공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="11" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">특<br>별<br>소<br>득<br>공<br>제</th>
                              <th class="left" rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">보험료</th>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">건강보험료<br>(노인장기요양보험료포함)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">고용보험료</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="16" style=" border-bottom:1px solid #e3e3e3;">특<br>별<br>세<br>액<br>공<br>제</th>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">보장성보험료</th>
                              <th >공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="7" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주택자금</th>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주택임차차입금원리금상환액</th>
                              <th class="left" colspan="2" style=" border-bottom:1px solid #e3e3e3;">대출기관</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="2" style=" border-bottom:1px solid #e3e3e3;">거주자</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">의료비</th>
                              <th >공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="5" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장기주택저당차입금이자상환액</th>
                              <th class="left" rowspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2011년 이전 차입분</th>
                              <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15년 미만</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15년~29년</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="3" style=" border-bottom:1px solid #e3e3e3;">교육비</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">30년이상</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2012년 이후 차입분<br>(15년이상)</th>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">고금리.비거치상환대출</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th rowspan="8" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기<br>부<br>금</th>
                              <th class="left" rowspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">정치자금기부금</th>
                              <th class="left" rowspan="2" style=" border-bottom:1px solid #e3e3e3;">10만원이하</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">그 밖의 대출</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기부금(이월분)</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" style=" border-bottom:1px solid #e3e3e3;">10만원초과</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">계</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="6" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">차감소득금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="2" style=" border-bottom:1px solid #e3e3e3;">법정기부금</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="13" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">그<br>밖<br>의<br>소<br>득<br>공<br>제</th>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">개인연금저축</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">소기업.소상공인 공제부금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="2" style=" border-bottom:1px solid #e3e3e3;">지정기부금</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" rowspan="3" colspan="2" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주택마련 저축 소득공제</th>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">청약저축</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">주택청약종합저축</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">계</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="3" style=" border-bottom:1px solid #e3e3e3;">근로자주택마련저축</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="4" style=" border-bottom:1px solid #e3e3e3;">표준세액공제</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">투자조합출자 등</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">납세조합공제</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">신용카드 등 사용액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">주택차입금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">우리사주조합 출연금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">외국납부</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">우리사주조합 기부금</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" rowspan="2" colspan="4" style=" border-bottom:1px solid #e3e3e3;">월세액</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">고용유지 중소기업 근로자</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">세액공제액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">목돈안드는 전세 이자상환액</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="5" style=" border-bottom:1px solid #e3e3e3;">세액 공제 계</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장기집합투자증권저축</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th colspan="5" style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right">&nbsp;</td>
						    </tr>
                            <tr>
                              <th class="left" colspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">그 밖의 소득공제 계</th>
                              <td class="right"><%=formatnumber(sum_oth_nhis,0)%>&nbsp;</td>
                              <th class="left" colspan="6" style=" border-bottom:1px solid #e3e3e3;">결정세액</th>
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
                                <td class="left" colspan="18" scope="col" style=" border-bottom:1px solid #e3e3e3;">소득·세액공제 명세(인적공제항목은 해당란에 ○표시(장애인 해당시는 해당코드 기재)를 하며, 각종 소득공제·세액공제 항목은 공제를 위하여 실제 지출한 금액을 적습니다)</td>
							</tr>
                            <tr>
								<td rowspan="3" style=" border-right:1px solid #e3e3e3; border-bottom:1px solid #FFFFFF;">&nbsp;</td>
                                <th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3;">인적공제항목</th>
                                <th colspan="11" scope="col" style=" border-bottom:1px solid #e3e3e3;">각종 소득공제 항목</th>
							</tr>
                            <tr>
                                <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">관계</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">성명</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">기본<br>공제</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">경로<br>우대</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">출산<br>입양</th>
                                <th rowspan="2" scope="col">자료구분</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                                <th rowspan="2" scope="col">의료비</th>
                                <th rowspan="2" scope="col">교육비</th>
                                <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">신용카드 등 사용액</th>
                                <th rowspan="2" scope="col">기부금</th>
							</tr>
                            <tr>
                                <th class="first" scope="col">내외</th>
                                <th scope="col">주민등록번호</th>
								<th scope="col">부녀자</th>
                                <th scope="col">한부모</th>
                                <th scope="col">장애인</th>
                                <th scope="col">6세이하</th>
                                <th scope="col">건강.고용등</th>
                                <th scope="col">보장성</th>
                                <th scope="col">신용카드<br>(전통시장·대중교통제외)</th>
                                <th scope="col">직불카드 등<br>(전통시장·대중교통제외)</th>
                                <th scope="col">현금영수증<br>(전통시장·대중교통제외)</th>
                                <th scope="col">전통시장<br>사용액</th>
                                <th scope="col">대중교통<br>이용액</th>
							</tr>
                            </thead>
                            <tbody>
                            <tr>
                                <td rowspan="20" style=" border-right:1px solid #e3e3e3;">Ⅰ. 인적공제 및 소득공제 명세서</td>
                                <td rowspan="2" colspan="2" style=" border-top:2px solid #515254;">인원 :&nbsp;<%=rs_year("y_support_cnt")%>&nbsp;(다자녀:&nbsp;<%=rs_year("y_daja_cnt")%>명)</td>
                                <td colspan="2" style=" border-top:2px solid #515254;"><%=y_basic_cnt%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=rs_year("y_old_cnt")%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=rs_year("y_holt_cnt")%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;">국세청자료</td>
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
                                <td>그밖의자료</td>
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
							if f_rel = "본인" then
							      f_person_no = "근로자 본인"
								  rel_chk = "0"
							   else
							      f_person_no = cstr(mid(rs("f_person_no"),1,6)) + "-" + cstr(mid(rs("f_person_no"),7,7))
								  if f_rel = "부" or f_rel = "모" or f_rel = "조부" or f_rel = "조모" then
								         rel_chk = "1"
								     elseif f_rel = "장인" or f_rel = "장모" or f_rel = "외조부" or f_rel = "외조모" then
									            rel_chk = "2"
											elseif f_rel = "남편" or f_rel = "아내" then
											           rel_chk = "3"
											elseif f_rel = "아들" or f_rel = "딸" then	
											           rel_chk = "4"
												   else
												       rel_chk = "5"
								  end if
						    end if
							if f_rel = "형(형제자매)" or f_rel = "매(형제자매)" or f_rel = "제(형제자매)" or f_rel = "자(형제자매)" then
							   rel_chk = "6"
							end if
							if f_pensioner = "Y" then
							   rel_chk = "7"
							end if
							if f_witak = "Y" then
							   rel_chk = "8"
							end if
							if rs("f_rel") = "본인" or rs("f_wife") = "Y" or rs("f_age20") = "Y" or rs("f_age60") = "Y" or rs("f_old") = "Y" then
							        basic_chk = "○"
							   else
							        basic_chk = ""
						    end if
							if rs("f_old") = "Y" then
							        old_chk = "○"
							   else
							        old_chk = ""
						    end if
							if rs("f_holt") = "Y" then
							        holt_chk = "○"
							   else
							        holt_chk = ""
						    end if
							if rs("f_woman") = "Y" then
							        woman_chk = "○"
							   else
							        woman_chk = ""
						    end if
							if rs("f_single") = "Y" then
							        single_chk = "○"
							   else
							        single_chk = ""
						    end if
							if rs("f_disab") = "Y" then
							        disab_chk = "○"
							   else
							        disab_chk = ""
						    end if
							if rs("f_children") = "Y" then
							        children_chk = "○"
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
                                <td style=" border-top:2px solid #515254;">국세청자료</td>
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
                                <td>그밖의자료</td>
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
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_wonchen_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_tax_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">원천징수영수증 출력</a>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

