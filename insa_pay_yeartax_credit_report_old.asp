<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim cc_tab(20,20)

'on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

inc_yyyy = cint(mid(now(),1,4)) - 1

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

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_cred = Server.CreateObject("ADODB.Recordset")
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
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
if not rs_year.eof then
       y_nhis_amt = rs_year("y_nhis_amt")
	   y_longcare_amt = rs_year("y_longcare_amt")
	   y_epi_amt = rs_year("y_epi_amt")
	   tot_pay = rs_year("y_total_pay") + rs_year("y_total_bonus") + rs_year("y_other_pay")
   else
       y_nhis_amt = 0
	   y_longcare_amt = 0
	   y_epi_amt = 0
end if

Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
'Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
	   tot_pay = tot_pay + rs_bef("b_pay") + rs_bef("b_bonus") + rs_bef("b_deem_bonus")
	rs_bef.MoveNext()
loop
rs_bef.close()

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

'sql = "select * from pay_yeartax_credit where c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' ORDER BY c_emp_no,c_person_no,c_id,c_seq ASC"
'rs_cred.Open Sql, Dbconn, 1

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
		   if rs_cred("c_id") = "�ſ�ī��" and rs_cred("c_market")  = "Y" then 
		         cc_tab(i,8) =  cc_tab(i,8) + rs_cred("c_nts_amt")   
				 cc_tab(i,9) =  cc_tab(i,9) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "�ſ�ī��" and rs_cred("c_transit")  = "Y" then 
		         cc_tab(i,10) =  cc_tab(i,10) + rs_cred("c_nts_amt")   
				 cc_tab(i,11) =  cc_tab(i,11) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "�ſ�ī��" and rs_cred("c_transit")  <> "Y" and rs_cred("c_transit") <> "Y" then 
		         cc_tab(i,6) =  cc_tab(i,6) + rs_cred("c_nts_amt")   
				 cc_tab(i,7) =  cc_tab(i,7) + rs_cred("c_other_amt")   
           end if
		   
		   if rs_cred("c_id") = "����ī��" and rs_cred("c_market")  = "Y" then 
		         cc_tab(i,14) =  cc_tab(i,14) + rs_cred("c_nts_amt")   
				 cc_tab(i,15) =  cc_tab(i,15) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "����ī��" and rs_cred("c_transit")  = "Y" then 
		         cc_tab(i,16) =  cc_tab(i,16) + rs_cred("c_nts_amt")   
				 cc_tab(i,17) =  cc_tab(i,17) + rs_cred("c_other_amt")   
           end if
		   if rs_cred("c_id") = "����ī��" and rs_cred("c_transit")  <> "Y" and rs_cred("c_transit") <> "Y" then 
		         cc_tab(i,12) =  cc_tab(i,12) + rs_cred("c_nts_amt")   
				 cc_tab(i,13) =  cc_tab(i,13) + rs_cred("c_other_amt")   
           end if
		   
		   if rs_cred("c_id") = "���ݿ�����" and rs_cred("c_market")  = "Y" then 
		         cc_tab(i,19) =  cc_tab(i,14) + rs_cred("c_nts_amt")   
           end if
		   if rs_cred("c_id") = "���ݿ�����" and rs_cred("c_transit")  = "Y" then 
		         cc_tab(i,20) =  cc_tab(i,16) + rs_cred("c_nts_amt")   
           end if
		   if rs_cred("c_id") = "���ݿ�����" and rs_cred("c_transit")  <> "Y" and rs_cred("c_transit") <> "Y" then 
		         cc_tab(i,18) =  cc_tab(i,12) + rs_cred("c_nts_amt")   
           end if
       end if
	next
	rs_cred.MoveNext()
loop
rs_cred.close()	

'sql = "select * from pay_yeartax_credit where c_year = '"&inc_yyyy&"' and c_emp_no = '"&emp_no&"' ORDER BY c_emp_no,c_person_no,c_id,c_seq ASC"
'Rs.Open Sql, Dbconn, 1


title_line = "��������-�ſ�ī��� �ҵ���� ��û��"
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
				<form action="insa_pay_yeartax_credit_report.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="20%" >
							<col width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td colspan="4">�ҵ��� ���� ����</td>
						    </tr>
                            <tr>
							  <th class="left" style=" border-top:1px solid #e3e3e3;">����</th>
                              <td><%=emp_name%></td>
                              <th class="left" style=" border-top:1px solid #e3e3e3;">�ֹε�Ϲ�ȣ(�Ǵ� �ܱ��ε�Ϲ�ȣ)</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th class="left">���θ�</th>
                              <td><%=company_name%></td>
                              <th class="left">��ü��</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
                              <td colspan="4">&nbsp;</td>
						    </tr>
						</thead>
					</table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="5%" >
                              <col width="5%" >
                              <col width="*" >
                              <col width="10%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="10%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1. ��������� �� �������ݾ� ��</th>
                              </tr>
                              <tr>
                                <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">���������</th>
                                <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">�ſ�ī��� ���ݾ�</th>
                              </tr>
                              <tr>
                                <th class="first" scope="col">1�����ܱ��α���</th>
                                <th scope="col">2����</th>
                                <th scope="col">3����</th>
                                <th scope="col">4�������</th>
                                <th scope="col">�ڷᱸ��</th>
                                <th scope="col">5�Ұ�<br>(6+7+8+9+10)</th>
                                <th scope="col">6�ſ�ī��<br>(������塤���߱��� ����)</th>
                                <th scope="col">7���ݿ�����<br>(������塤���߱��� ����)</th>
                                <th scope="col">8���ҡ�����ī��<br>(������塤���߱��� ����)</th>
                                <th scope="col">9����������<br>(�ſ�ī��,���ҡ�����ī��,���ݿ�����)</th>
                                <th scope="col">10���߱���<br>(�ſ�ī��,���ҡ�����ī��,���ݿ�����)</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						sum_market = 0
						sum_transit = 0
						sum_credit = 0
						sum_cash = 0
						sum_direct = 0
						sum_hap = 0
						
						for i = 1 to 20
                            if cc_tab(i,2) = "" or isnull(cc_tab(i,2)) then 
			                        exit for
		                       else 
							 c_year = cc_tab(i,1)
							 c_emp_no = cc_tab(i,2)
							 c_person_no = cc_tab(i,3)
							 
							 sql = "select * from pay_yeartax_family where f_year = '"&c_year&"' and f_emp_no = '"&c_emp_no&"' and f_person_no = '"&c_person_no&"'"
                             rs_fami.Open Sql, Dbconn, 1
                             if not rs_fami.eof then
							        f_national = rs_fami("f_national")
							        f_pensioner = rs_fami("f_pensioner")
							        f_witak = rs_fami("f_witak")
									f_birthday = rs_fami("f_birthday")
							    else
								    f_national = ""
							        f_pensioner = ""
							        f_witak = ""
									f_birthday = ""
							 end if
						     rs_fami.close()	
							 
							 c_rel = cc_tab(i,4)

							 nts_market = cc_tab(i,8) + cc_tab(i,14) + cc_tab(i,19) 
							 nts_transit = cc_tab(i,10) + cc_tab(i,16) + cc_tab(i,20)
							 other_market = cc_tab(i,9) + cc_tab(i,15)
							 other_transit = cc_tab(i,11) + cc_tab(i,17)
							 nts_hap = cc_tab(i,6) + cc_tab(i,12) + cc_tab(i,18) + nts_market + nts_transit
							 other_hap = cc_tab(i,7) + cc_tab(i,13) + other_market + other_transit
							 
							 sum_market = sum_market + nts_market + other_market
							 sum_transit = sum_transit + nts_transit + other_transit
							 sum_credit = sum_credit + cc_tab(i,6) + cc_tab(i,7)
							 sum_cash = sum_cash + cc_tab(i,18)
							 sum_direct =  sum_direct + cc_tab(i,12) + cc_tab(i,13)
							 sum_hap =  sum_hap + nts_hap + other_hap
	           			%>
							<tr>
                                <td rowspan="2"><%=f_national%>&nbsp;</td>
                                <td rowspan="2"><%=c_rel%>&nbsp;</td>
                                <td rowspan="2"><%=cc_tab(i,5)%>&nbsp;</td>
                                <td rowspan="2"><%=c_person_no%>&nbsp;</td>
                                <td class="left">����û �ڷ�</td>
                                <td class="right"><%=formatnumber(nts_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cc_tab(i,6),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cc_tab(i,18),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cc_tab(i,12),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(nts_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(nts_transit,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td class="left" style=" border-left:1px solid #e3e3e3;">�� ���� �ڷ�</td>
                                <td class="right"><%=formatnumber(other_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cc_tab(i,7),0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(cc_tab(i,13),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(other_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(other_transit,0)%>&nbsp;</td>
							</tr>
						<%
						    end if
						next
						%>
                        	<tr>
                                <td colspan="5">10�հ��</td>
                                <td class="right"><%=formatnumber(sum_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_credit,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_cash,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_direct,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_market,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_transit,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="11">&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">2. �ſ�ī��� �ҵ�������� ���</th>
                            </tr>
                        <%
						    tax15_31 = 0
							tax15_32 = 0
							market_tax = int(sum_market * (30 / 100))
							transit_tax = int(sum_transit * (30 / 100))
							cash_tax = int((sum_cash + sum_direct) * (30 / 100))
							credit_tax = int(sum_credit * (15 / 100))
							pay_tax = int(tot_pay * (25 / 100))
							if pay_tax <= sum_credit then 
							       tax15_31 = int(pay_tax * (15 / 100))
								   tax15_3 = tax15_31
							   else
							       tax15_32 = int((sum_credit * (15 / 100)) + ((pay_tax - sum_credit) * (30 / 100)))
								   tax15_3 = tax15_32
							end if
						%>
                            <tr>
                                <th rowspan="2" colspan="3" style="background:#f8f8f8;">11����������<br>������<br>(9*30%)</th>
                                <th rowspan="2" colspan="2" style="background:#f8f8f8;">12���߱����̿��<br>������<br>(10*30%)</th>
                                <th rowspan="2" style="background:#f8f8f8;">13���ҡ�����ī��<br>���ݿ������� ����<br>(7+8)*30%</th>
                                <th rowspan="2" style="background:#f8f8f8;">14�ſ�ī�����<br>������<br>(6*15%)</th>
                                <th colspan="3" style="background:#f8f8f8;">15�������ܱݾ� ���</th>
                                <th rowspan="2" style="background:#f8f8f8;">16üũī���<br>����������<br>������</th>
							</tr>
                            <tr>
                                <th style="background:#f8f8f8; border-left:1px solid #e3e3e3;"">15-1<br>�ѱ޿�</th>
                                <th style="background:#f8f8f8;">15-2<br>�������ݾ�<br>(15-1*25%)</th>
                                <th style="background:#f8f8f8;">15-3<br>�������ܱݾ�</th>
							</tr>
                            <tr>
                                <td colspan="3" class="right"><%=formatnumber(market_tax,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(transit_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cash_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(credit_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(tot_pay,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pay_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(tax15_3,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="3" style="background:#f8f8f8;">17�������ɱݾ�<br>[11+12+13+14-(15-3)+16]<br>(9*30%)</th>
                                <th colspan="2" style="background:#f8f8f8;">18�����ѵ���<br>[3�鸸����<br>(15-1)*20%�� �����ݾ�]</th>
                                <th style="background:#f8f8f8;">19�Ϲ� �����ݾ�<br>(17�� 18�� �����ݾ�)</th>
                                <th colspan="2" style="background:#f8f8f8;">20������� �߰������ݾ�<br>[17-18(�����̸� 0���� ��)��<br>11�� �����ݾ�(�ѵ�:1�鸸��)]</th>
                                <th colspan="2" style="background:#f8f8f8;">21���߱��� �߰������ݾ�<br>[17-20-19(�����̸� 0���� ��)��<br>12�� �����ݾ�(�ѵ�:1�鸸��)]</th>
                                <th style="background:#f8f8f8;">22���� �����ݾ�<br>[19+20+21]</th>
							</tr>
                            <tr>
                                <td colspan="3" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">15-3 ���</th>
                            </tr>
                            <tr>
                                <th colspan="4" style="background:#f8f8f8;">����</th>
                                <th colspan="5" style="background:#f8f8f8;">����</th>
                                <th colspan="2" style="background:#f8f8f8;">15-3</th>
							</tr>
                            <tr>
                                <th colspan="4" style="background:#f8f8f8;">15-2 �������ݾ� �� �ſ�ī�����6</th>
                                <td colspan="5" class="left" >15-2 * 15%</td>
                                <td colspan="2" class="right"><%=formatnumber(tax15_31,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="4" style="background:#f8f8f8;">15-2 �������ݾ� > �ſ�ī�����6</th>
                                <td colspan="5" class="left" >6 * 15% + [(15-2) - 6] * 30%</td>
                                <td colspan="2" class="right"><%=formatnumber(tax15_32,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">16 ���</th>
                            </tr>
                            <tr>
                                <th colspan="3" style="background:#f8f8f8;">����</th>
                                <th style="background:#f8f8f8;">�����Ⱓ</th>
                                <th colspan="2" style="background:#f8f8f8;">�ݾ�</th>
                                <th colspan="5" class="left" style="background:#f8f8f8;">16üũī�� �� ���� ������ ������</th>
							</tr>
                            <tr>
                                <th rowspan="2" colspan="3" style="background:#f8f8f8;">������ �ſ�ī��� ����</th>
                                <th style="background:#f8f8f8;">2013��</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td rowspan="2" colspan="5" class="left" >&nbsp;</td>
							</tr>
                            <tr>
                                <th style=" border-left:1px solid #e3e3e3; background:#f8f8f8;">2014��</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th rowspan="2" colspan="3" style="background:#f8f8f8;">������ �ſ�ī��� ����</th>
                                <th style="background:#f8f8f8;">2013��</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td rowspan="2" colspan="5" class="left" >(����)<br>��2013�� ������ �ſ�ī�� �� ���� �� 2014�� ������ �ſ�ī�� �� ���� : "0"<br>��2013�� ������ �ſ�ī�� �� ���� < 2014�� ������ �ſ�ī���<br>����:(2014�� �Ϲݱ� �߰����������� - 2013�� �߰�����������<br>*50%) * 10%(��, ������ ��� "0")</td>
							</tr>
                            <tr>
                                <th style="border-left:1px solid #e3e3e3; background:#f8f8f8;">2014��<br>�Ϲݱ�</th>
                                <td colspan="2" class="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="11">&nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="11" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">������Ư�����ѹ� ����ɡ���121����2��8�׿� ���� �ſ�ī�� �� ���ݾ׿� ���� �ҵ������ ��û �մϴ�.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 �� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��û�� : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(���� �Ǵ� ��)<br></td>
                            </tr>
                            <tr>
                                <td colspan="11">&nbsp;</td>
                            </tr>
                            <tr>
                                <th style="background:#f8f8f8;">���񼭷�</th>
                                <td colspan="9" class="left" >�ſ�ī�� �� ���ݾ� Ȯ�μ�(���� ��74ȣ��5������ ���մϴ�) �Ǵ� ����û Ȩ���������� �����ϴ� �ſ�ī�� �� ���ݾ� ���� ����� ���� 1��</td>
                                <td>������ ����</td>
							</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType04">�Ƿ�����޸���</a>
                    <a href="insa_pay_yeartax_donation_report.asp" class="btnType04">��αݸ���</a>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_credit_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_credit_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">�ſ�ī��� ���� ���</a>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

