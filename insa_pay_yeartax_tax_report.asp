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
		y_woman = "○"
	else
		y_woman = ""
	end If

	if rs_year("y_single") = "Y" then
		y_single = "○"
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
'	y_woman = "○"
'else
'	y_woman = ""
'end If

'if rs_year("y_single") = "Y" then
'	y_single = "○"
'else
'	y_single = ""
'end if

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

title_line = "소득·세액 공제신고서/근로자 소득·세액 공제신고서(2014년 소득에 대한 연말정산용)"
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
                              <td colspan="7" class="left">※ 근로소득자는 신고서에 소득·세액 공제 증명서를 첨부하여 원천징수의무자(소속회사 등)에게 제출하며, 원천징수의무자는 신고서 및 첨부서류를 확인하여 근로소득 세액계산을 하고 근로소득자에게 즉시<br>근로소득원천징수영수증을 발급해야 합니다. 연말정산 시 근로소득자에게 환급이 발생하는 경우 원천징수의무자는 근로소득자에게 환급세액을 지급해야 합니다.</td>
						    </tr>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">성명</th>
                              <td style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%></td>
                              <th style=" border-top:1px solid #e3e3e3;">인사코드</th>
                              <td colspan="2"><%=emp_no%></td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">주민등록번호(또는 외국인등록번호)</th>
                              <td ><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">법인명</th>
                              <td colspan="3"><%=company_name%></td>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">업체명</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">세대주 여부</th>
                              <td colspan="3"><%=householder%></td>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">국적</th>
                              <td>(국적코드:<%=y_national_code%> )&nbsp;<%=y_national%></td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">근무기간</th>
                              <td colspan="3"><%=y_from_date%>&nbsp;∼&nbsp;<%=y_to_date%></td>
                              <th class="left" style=" border-bottom:1px solid #e3e3e3;">감면기간</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">거주구분</th>
                              <td colspan="3"><%=yy_live%></td>
                              <th class="left">거주지국</th>
                              <td><%=y_national%>&nbsp;(국적코드:<%=y_national_code%>)</td>
						    </tr>
                            <tr>
							  <th colspan="2">인적공제 항목 변동 여부</th>
                              <td colspan="3"><%=yy_change%></td>
                              <td colspan="2" class="left" style="color:#ff0000;">※ 인적공제 항목이 전년과 동일한 경우에도 주민등록표등본을 제출해주시기 바랍니다.</td>
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
							  <th colspan="2" style=" border-bottom:2px solid #515254;">구분</th>
                              <th colspan="3" style=" border-bottom:2px solid #515254;">지출명세</th>
                              <th style=" border-bottom:2px solid #515254;">지출구분</th>
                              <th style=" border-bottom:2px solid #515254;">금액</th>
                              <th style=" border-bottom:2px solid #515254;">한도액</th>
                              <th style=" border-bottom:2px solid #515254;">공제액</th>
						    </tr>
                            <tr>
							  <th rowspan="5">Ⅱ.연금보험료공제</th>
                              <th rowspan="5">연금보험료<br>(국민연금,공무원연금,군인연금,교직원연금,연금계좌등)</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3;">국민연금보험료</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(b_nps,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(b_nps_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right" ><%=formatnumber(y_nps_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nps_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">국민연금보험료 외의 공적연금보험료</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(y_nps_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nps_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(y_nps_other,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nps_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3;">연금보험료 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(a_tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(a_tot_tax,0)%>&nbsp;</td>
						    </tr>


                            <tr>
							  <th rowspan="13" style=" border-top:2px solid #515254;">Ⅲ. 특별소득공제</th>
                              <th rowspan="5" style=" border-top:2px solid #515254;">보험료</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">국민건강보험<br>(노인장기요양보험 포함)</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">보험료</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(b_nhis,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">전액</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(b_nhis_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;"">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right" ><%=formatnumber(y_nhis_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_nhis_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">고용보험</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">종(전)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(b_epi,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(b_epi_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주(현)근무지</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                              <td class="right"><%=formatnumber(y_epi_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">전액</th>
                              <td class="right"><%=formatnumber(y_epi_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3;">보험료 계</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(i_tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(i_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="8" style=" border-top:2px solid #515254; border-left:1px solid #e3e3e3;">주택자금</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">주택임차차입금</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">대출기관차입</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">원리금상환액</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(h_lender_amt,0)%>&nbsp;</td>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">작성방법 참조</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(h_lender_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">거주자 차입</th>
                              <td class="right"><%=formatnumber(h_person_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_person_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="5" style=" border-left:1px solid #e3e3e3;">장기주택저당차입금</th>
                              <th rowspan="3" style=" border-bottom:1px solid #e3e3e3;">2011년 이전<br>차입분</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">15년미만</th>
                              <th rowspan="5" >이자상환액</th>
                              <td class="right"><%=formatnumber(h_long15_amt,0)%>&nbsp;</td>
                              <th rowspan="5" style=" border-bottom:1px solid #e3e3e3;">작성방법 참조</th>
                              <td class="right"><%=formatnumber(h_long15_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15년 ~ 29년</th>
                              <td class="right"><%=formatnumber(h_long29_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_long29_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">30년</th>
                              <td class="right"><%=formatnumber(h_long30_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_long30_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th rowspan="2" style=" border-left:1px solid #e3e3e3; ">2012년 이후<br>차입분<br>(15년이상)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">고정금리.비거치상환대출</th>
                              <td class="right"><%=formatnumber(h_fixed_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_fixed_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; ">기타대출</th>
                              <td class="right"><%=formatnumber(h_other_amt,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(h_other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-top:1px solid #e3e3e3;">주택자금 공제액 계</th>
                              <th style=" border-top:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(h_tot_amt,0)%>&nbsp;</td>
                              <th >&nbsp;</th>
                              <td class="right"><%=formatnumber(h_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="21" style=" border-top:2px solid #515254;">Ⅳ. 그 밖 의 소 득 공 제</th>
                              <th colspan="4" style=" border-top:2px solid #515254; border-bottom:1px solid #e3e3e3;">개인연금저축(2000년 12월 31일 이전 가입)</th>
                              <th style=" border-top:2px solid #515254; border-bottom:1px solid #e3e3e3;">납입금액</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(tot_2000,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-top:2px solid #515254;">불입액40%와(72만원)</th>
                              <td class="right" style=" border-top:2px solid #515254;"><%=formatnumber(tax_2000,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">소기업·소상공인 공제부금</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >납입금액</th>
                              <td class="right" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주택마련저축</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3; ">청약저축</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">납입금액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_cheng,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">근로자주택마련저축</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">납입금액</th>
                              <td class="right" ><%=formatnumber(tot_gunro,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_gunro,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">주택청약종합저축</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">납입금액</th>
                              <td class="right" ><%=formatnumber(tot_jutak,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_jutak,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">주택마련저축 소득공제 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">투자조합 출자등</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3; ">2012년 출자·투자분</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">출자·투자금액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_cheng,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">2013년 출자·투자분</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">출자·투자금액</th>
                              <td class="right" ><%=formatnumber(tot_gunro,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_gunro,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">2014년 이후 출자·투자분</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">출자·투자금액</th>
                              <td class="right" ><%=formatnumber(tot_jutak,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(tax_jutak,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">투자조합 출자 등 소득공제 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>



                            <tr>
							  <th rowspan="10" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">신용카드 등 사용액</th>
                              <th colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3; ">①신용카드(전통시장·대중교통사용분제외)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">②직불·선불카드(전통시장·대중교통사용분제외)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left"style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">③현금영수증(전통시장·대중교통사용분제외)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">④전통시장사용분</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">⑤대중교통사용분</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">계(①+②+③+④+⑤</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">⑥본인 신용카드등 사용액(2013년)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">⑦본인 신용카드등 사용액(2014년)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">⑧본인 추가공제율 사용분(2013년)<br>-직불,선불,현금영수증,전통시장,대중교통</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th colspan="3" class="left" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">⑨본인 추가공제율 사용분(2014년하반기)<br>-직불,선불,현금영수증,전통시장,대중교통</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <th class="right" >&nbsp;</th>
						    </tr>


                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">우리사주조합 출연금</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">출연금액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">우리사주조합 기부금</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">기부금액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">고용유지중소기업 근로자</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">임금삭감액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">목돈 안 드는 전세 이자상환액</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">이자상환액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
                              <td class="right" ><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; ">장기집합투자증권저축</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">납입금액</th>
                              <td class="right" ><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">작성방법 참조</th>
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
							  <th colspan="2" style=" border-bottom:2px solid #515254;">구분</th>
                              <th colspan="4" style=" border-bottom:2px solid #515254;">세액감면·공제명세</th>
                              <th colspan="6" style=" border-bottom:2px solid #515254;">세액감면·공제명세</th>
						    </tr>
                            <tr>
							  <th rowspan="35" >Ⅴ. 세액감면및 공제</th>
                              <th rowspan="5" style="border-bottom:1px solid #e3e3e3;">세액감면</th>
                              <th rowspan="4" style="border-bottom:1px solid #e3e3e3;">외국인근로자</th>
                              <th colspan="2" style="border-bottom:1px solid #e3e3e3;">입국목적</th>
                              <td colspan="7">[ ]정부간 협약 [ ]기술도입계약 [ ]「조세특례제한법」상 감면 [ ]조세조약 상 감면</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기술도입계약 또는 근로제공일</th>
                              <td ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; " >감면기간 만료일</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">외국인 근로소득에 대한 감면</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >접수일</th>
                              <td colspan="2" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >제출일</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">근로소득에 대한 조세조약 상 면제</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >접수일</th>
                              <td colspan="2" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >제출일</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">중소기업 취업자 감면</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >취업일</th>
                              <td colspan="2" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >감면기간 종료일</th>
                              <td colspan="3" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="30" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">세액공제</th>
                              <th colspan="4" style="border-bottom:1px solid #e3e3e3;">공 제 종 류</th>
                              <th colspan="2" style="border-bottom:1px solid #e3e3e3;">명세</th>
                              <th style="border-bottom:1px solid #e3e3e3;">한도액</th>
                              <th style="border-bottom:1px solid #e3e3e3;">공제대상금액</th>
                              <th style="border-bottom:1px solid #e3e3e3;">공제율</th>
                              <th style="border-bottom:1px solid #e3e3e3;">공제세액</th>
						    </tr>
                            <tr>
                              <th rowspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">연금계좌</th>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">과학기술인공제</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >납입금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th rowspan="3" style=" border-bottom:1px solid #e3e3e3; " >작성방법 참조</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="4">12%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">「근로자퇴직급여 보장법」에 따른 퇴직연금</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >납입금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">연금저축</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >납입금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">연금계좌 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right" ><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>

                            <tr>
                              <th rowspan="17" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">특별세액공제</th>
                              <th rowspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">보험료</th>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">보장성</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >보험료</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >100만원</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="3">12%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장애인전용보장성</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >보험료</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >100만원</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">보험료 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">의료비</th>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">본인·65세이상자·장애인</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >지출액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >작성방법 참조</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="3">15%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">그 밖의 공제대상자</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >지출액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >작성방법 참조</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">의료비 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="6" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">교육비</th>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">소득자 본인</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >공납금(대학원포함)</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >전액</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td rowspan="6">15%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">취학전 아동(<%=adong_cnt%> 명)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >유치원·학원비등</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >1명당 300만원</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">초·중·고등학교(<%=adong_cnt%> 명)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >공납금</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >1명당 300만원</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">대학생(대학원 불포함)(<%=adong_cnt%> 명)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >공납금</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >1명당 300만원</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">장애인(<%=adong_cnt%> 명)</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >특수교육비</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >전액</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">교육비 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right" ><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>


                            <tr>
                              <th rowspan="5" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기부금</th>
                              <th rowspan="2" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">정치자금기부금</th>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">10만원이하</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >기부금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th rowspan="4" style=" border-bottom:1px solid #e3e3e3; " >작성방법 참조</th>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >100/110</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">10만원초과</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >기부금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >15%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">법정기부금</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >기부금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >25%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">지정기부금</th>
                              <th style=" border-bottom:1px solid #e3e3e3; " >기부금액</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td >25%</td>
                              <td class="right"><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기부금 계</th>
                              <th style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</th>
                              <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th >&nbsp;</th>
                              <td class="right" style=" border-bottom:1px solid #e3e3e3; "><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
                              <td style=" border-bottom:1px solid #e3e3e3; ">&nbsp;</td>
                              <td class="right" style=" border-bottom:1px solid #e3e3e3; "><%=formatnumber(tax_sosang,0)%>&nbsp;</td>
						    </tr>

                            <tr>
                              <th rowspan="6" colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">외국납부세액</th>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">국외원천소득</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >납세액(외화)</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >납세액(원화)</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >&nbsp;</th>
                              <th colspan="2" >&nbsp;</th>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >납세국명</th>
                              <td colspan="2" class="right"><%=de_tax_nation%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >납부일</th>
                              <td colspan="2" class="right"><%=de_tax_date%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >신청서제출일</th>
                              <td colspan="2" class="right"><%=de_tax_nation%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >국외근무처</th>
                              <td colspan="2" class="right"><%=de_tax_date%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >근무기간</th>
                              <td colspan="2" class="right"><%=de_tax_nation%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >직책</th>
                              <td colspan="2" class="right"><%=de_tax_date%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">주택자금차입금이자세액공제</th>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >이자상환액</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >30%</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="4" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">월세세액공제</th>
                              <th style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >월세액</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3; " >10%</th>
                              <td colspan="2" class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
						    </tr>
                            <tr>
								<td colspan="12" scope="col" class="left" style="border-top:2px solid #515254;">신고인은 「소득세법」 제140조에 따라 위의 내용을 신고하며,<br>위 내용을 충분히 검토하였고 신고인이 알고 있는 사실 그래로를 정확하게 적었음을 확인합니다.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 년 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;월 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;일<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;신청인 : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(서명 또는 인)<br></td></td>
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
								<td colspan="6" scope="col" class="left">Ⅵ. 추가 제출 서류</td>
							</tr>
                            <tr>
								<td colspan="5" scope="col" class="left" >1. 외국인근로자 단일세율적용신청서 제출 여부(○ 또는 ×로 적습니다)</td>
                                <td scope="col" >제출 (&nbsp;&nbsp;&nbsp;)</td>
							</tr>
                            <tr>
								<th rowspan="2" scope="col" style=" border-top:1px solid #e3e3e3; "  >2. 종(전)근무지 명세</th>
                                <th scope="col" style="border-top:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >종(전)근무지명</th>
                                <td scope="col" ><%=de_tax_date%>&nbsp;</td>
                                <th scope="col" style="border-top:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3; " >종(전)급여총액</th>
                                <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
                                <td rowspan="2" scope="col" >종(전)근무지 근로소득<br>원천징수영수증 제출(&nbsp;&nbsp;)</td>
							</tr>
                            <tr>
                                <th scope="col" style=" border-left:1px solid #e3e3e3; " >사업자등록번호</th>
                                <td scope="col" ><%=de_tax_date%>&nbsp;</td>
                                <th scope="col" >종(전)결정세액</th>
                                <td class="right"><%=formatnumber(o_sosang,0)%>&nbsp;</td>
							</tr>
                            <tr>
								<td colspan="3" scope="col" class="left">3. 연금·저축 등 소득·세액 공제명세서 제출여부(○ 또는 ×로 적습니다)</td>
                                <td colspan="3" scope="col" class="left">제출(&nbsp;&nbsp;) ※ 연금계좌, 주택마련저축 등 소득·세액공제를 신청한 경우 해당 명세서를 제출해야 합니다.</td>
							</tr>
                            <tr>
								<td colspan="3" scope="col" class="left">4. 월세액·거주자 간 주택임차차임금 원리금상환액 소득공제 명세서 제출여부(○ 또는 ×로 적습니다)</td>
                                <td colspan="3" scope="col" class="left">제출(&nbsp;&nbsp;) ※ 월세액·거주자 간 주택임차차임금 원리금상환액 소득공제를 신청한 경우 해당 명세서를 제출해야 합니다.</td>
							</tr>
                            <tr>
								<td colspan="2" scope="col" class="left">5. 그 밖의 추가 제출 서류</td>
                                <td colspan="4" scope="col" >①의료비지급명세서(&nbsp;&nbsp;), ②기부금명세서(&nbsp;&nbsp;), ③소득공제 증명서류(&nbsp;&nbsp;)</td>
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
								<td scope="col" style=" border-bottom:2px solid #515254;">유의사항</td>
							</tr>
                            <tr>
								<td scope="col" class="left" >1. 근로소득자가 종(전)근무지 근로소득을 원천징수의무자에게 신고하지 않은 경우에는 근로소득자 본인이 종합소득세 신고를 해야 하며, 신고하지 않은 경우 가산세 부과 등 불이익이 따릅니다.<br><br>2. 현 근무지의 연금보험료·국민건강보험료 및 고용보험료 등은 신고인이 작성하지 않아도 됩니다.<br><br>3. 공제금액란은 근로자가 원천징수의무자에게 제출하는 경우 적지 않을 수 있습니다.</td>
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
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_tax_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_tax_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">소득공제신고서 출력</a>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>

