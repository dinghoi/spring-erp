<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

emp_no=Request("emp_no")
emp_name=Request("emp_name")
inc_yyyy=Request("inc_yyyy")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

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

tot_cnt = 0
tot_amt = 0

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
rs_medi.Open Sql, Dbconn, 1
'Set rs_medi = DbConn.Execute(SQL)
do until rs_medi.eof
         tot_cnt = tot_cnt + int(rs_medi("m_cnt"))	
		 tot_amt = tot_amt + int(rs_medi("m_amt"))
	rs_medi.MoveNext()
loop
rs_medi.close()	

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�ſ�ī��� �ҵ���� ��û��"
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
			function goAction () {
		  		 window.close () ;
			}
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //�Ӹ��� ����
                factory.printing.footer = ""; //������ ����
                factory.printing.portrait = true; //��¹��� ����: true - ����, false - ����
                factory.printing.leftMargin = 13; //���� ���� ����
                factory.printing.topMargin = 10; //���� ���� ����
                factory.printing.rightMargin = 13; //�����P ���� ����
                factory.printing.bottomMargin = 15; //�ٴ� ���� ����
        //		factory.printing.SetMarginMeasure(2); //�׵θ� ���� ������ ������ ��ġ�� ����
        //		factory.printing.printer = ""; //������ �� ������ �̸�
        //		factory.printing.paperSize = "A4"; //��������
        //		factory.printing.pageSource = "Manusal feed"; //���� �ǵ� ���
        //		factory.printing.collate = true; //������� ����ϱ�
        //		factory.printing.copies = "1"; //�μ��� �ż�
        //		factory.printing.SetPageRange(true,1,1); //true�� �����ϰ� 1,3�̸� 1���� 3������ ���
        //		factory.printing.Printer(true); //����ϱ�
                factory.printing.Preview(); //�����츦 ���ؼ� ���
                factory.printing.Print(false); //�����츦 ���ؼ� ���
            }
        </script>
    <style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
    </style>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="wrap">			
			<div id="container">
				<form action="insa_pay_yeartax_credit_print.asp" method="post" name="frm">
				<div class="gView">
				<table width="1150" cellpadding="0" cellspacing="0">
                   <tr>
                      <td class="style20C"><%=title_line%></td>
                   </tr>
                   <tr>
                      <td height="20" class="style20C">&nbsp;</td>
                   </tr>
                </table>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
				        <colgroup>
							<col height="30px" width="20%" >
							<col height="30px" width="30%" >
							<col height="30px" width="20%" >
							<col height="30px" width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td colspan="4" height="30" align="center" class="style12C">�ҵ��� ���� ����</td>
						    </tr>
                            <tr>
							  <th height="30" align="left" style=" border-top:1px solid #e3e3e3;">�缺��</th>
                              <td height="30" align="center"><%=emp_name%></td>
                              <th height="30" align="left" style=" border-top:1px solid #e3e3e3;">���ֹε�Ϲ�ȣ(�Ǵ� �ܱ��ε�Ϲ�ȣ)</th>
                              <td height="30" align="center"><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th height="30" align="left">����θ�</th>
                              <td height="30" align="center"><%=company_name%></td>
                              <th height="30" align="left">���ü��</th>
                              <td height="30" align="center">&nbsp;</td>
						    </tr>
                            <tr>
                              <td colspan="4" height="20" align="center" class="style12C">&nbsp;</td>
						    </tr>
						</thead>
				  </table>
					<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
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
                                <th colspan="11" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1. ��������� �� �������ݾ� ��</th>
                              </tr>
                              <tr>
                                <th colspan="4" height="30" align="center" scope="col" style=" border-bottom:1px solid #e3e3e3;">���������</th>
                                <th colspan="7" height="30" align="center" scope="col" style=" border-bottom:1px solid #e3e3e3;">�ſ�ī��� ���ݾ�</th>
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
						
						do until rs.eof
						
						   chk_sum = Rs("c_credit_nts") + Rs("c_credit_other") + Rs("c_cash_nts") + Rs("c_direct_nts") + Rs("c_direct_other") + Rs("c_market_nts") + Rs("c_market_other") + Rs("c_transit_nts") + Rs("c_transit_other")
						
                          if chk_sum > 0 then 
                             f_national = Rs("f_national") 						 

							 nts_market = Rs("c_market_nts")
							 nts_transit = Rs("c_transit_nts")
							 other_market = Rs("c_market_other")
							 other_transit = Rs("c_transit_other")
							 nts_hap = Rs("c_credit_nts") + Rs("c_cash_nts") + Rs("c_direct_nts") + nts_market + nts_transit
							 other_hap = Rs("c_credit_other") + Rs("c_direct_other") + other_market + other_transit
							 
							 sum_market = sum_market + nts_market + other_market
							 sum_transit = sum_transit + nts_transit + other_transit
							 sum_credit = sum_credit + Rs("c_credit_nts") + Rs("c_credit_other")
							 sum_cash = sum_cash + Rs("c_cash_nts")
							 sum_direct =  sum_direct + Rs("c_direct_nts") + Rs("c_direct_other")
							 sum_hap =  sum_hap + nts_hap + other_hap
	           			%>
							<tr>
                                <td rowspan="2" height="30" align="center"><%=f_national%>&nbsp;</td>
                                <td rowspan="2" align="center"><%=Rs("f_rel")%>&nbsp;</td>
                                <td rowspan="2" align="center"><%=Rs("f_family_name")%>&nbsp;</td>
                                <td rowspan="2" align="center"><%=Rs("f_person_no")%>&nbsp;</td>
                                <td align="left">����û �ڷ�</td>
                                <td align="right"><%=formatnumber(nts_hap,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs("c_credit_nts"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs("c_cash_nts"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs("c_direct_nts"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(nts_market,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(nts_transit,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td class="first" height="30" align="center" style=" border-left:1px solid #e3e3e3;">�� ���� �ڷ�</td>
                                <td align="right"><%=formatnumber(other_hap,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs("c_credit_other"),0)%>&nbsp;</td>
                                <td align="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs("c_direct_other"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(other_market,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(other_transit,0)%>&nbsp;</td>
							</tr>
						<%
						    end if							
							rs.movenext()
						loop
						rs.close()
						    if sum_hap > 0 then						
						%>
                        	<tr>
                                <td colspan="5" height="30" align="center">10�հ��</td>
                                <td align="right"><%=formatnumber(sum_hap,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(sum_credit,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(sum_cash,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(sum_direct,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(sum_market,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(sum_transit,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="11" height="20" align="center">&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="11" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">2. �ſ�ī��� �ҵ�������� ���</th>
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
                                <th rowspan="2" colspan="3" height="30" align="center" style="background:#f8f8f8;">11����������<br>������<br>(9*30%)</th>
                                <th rowspan="2" colspan="2" style="background:#f8f8f8;">12���߱����̿��<br>������<br>(10*30%)</th>
                                <th rowspan="2" align="center" style="background:#f8f8f8;">13���ҡ�����ī��<br>���ݿ������� ����<br>(7+8)*30%</th>
                                <th rowspan="2" align="center" style="background:#f8f8f8;">14�ſ�ī�����<br>������<br>(6*15%)</th>
                                <th colspan="3" align="center" style="background:#f8f8f8;">15�������ܱݾ� ���</th>
                                <th rowspan="2" align="center" style="background:#f8f8f8;">16üũī���<br>����������<br>������</th>
							</tr>
                            <tr>
                                <th height="30" align="center" style="background:#f8f8f8; border-left:1px solid #e3e3e3;"">15-1<br>�ѱ޿�</th>
                                <th align="center" style="background:#f8f8f8;">15-2<br>�������ݾ�<br>(15-1*25%)</th>
                                <th align="center" style="background:#f8f8f8;">15-3<br>�������ܱݾ�</th>
							</tr>
                            <tr>
                                <td colspan="3" height="30" align="right" ><%=formatnumber(market_tax,0)%>&nbsp;</td>
                                <td colspan="2" align="right"><%=formatnumber(transit_tax,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(cash_tax,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(credit_tax,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(tot_pay,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(pay_tax,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(tax15_3,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="3" height="30" align="center" style="background:#f8f8f8;">17�������ɱݾ�<br>[11+12+13+14-(15-3)+16]<br>(9*30%)</th>
                                <th colspan="2" align="center"style="background:#f8f8f8;">18�����ѵ���<br>[3�鸸����<br>(15-1)*20%�� �����ݾ�]</th>
                                <th align="center" style="background:#f8f8f8;">19�Ϲ� �����ݾ�<br>(17�� 18�� �����ݾ�)</th>
                                <th colspan="2" style="background:#f8f8f8;">20������� �߰������ݾ�<br>[17-18(�����̸� 0���� ��)��<br>11�� �����ݾ�(�ѵ�:1�鸸��)]</th>
                                <th colspan="2" align="center" style="background:#f8f8f8;">21���߱��� �߰������ݾ�<br>[17-20-19(�����̸� 0���� ��)��<br>12�� �����ݾ�(�ѵ�:1�鸸��)]</th>
                                <th align="center" style="background:#f8f8f8;">22���� �����ݾ�<br>[19+20+21]</th>
							</tr>
                            <tr>
                                <td colspan="3" height="30" align="right" ><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="11" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">15-3 ���</th>
                            </tr>
                            <tr>
                                <th colspan="4" height="30" align="center" style="background:#f8f8f8;">����</th>
                                <th colspan="5" align="center"style="background:#f8f8f8;">����</th>
                                <th colspan="2" align="center" style="background:#f8f8f8;">15-3</th>
							</tr>
                            <tr>
                                <th colspan="4" height="30" align="center" style="background:#f8f8f8;">15-2 �������ݾ� �� �ſ�ī�����6</th>
                                <td colspan="5" align="left" >15-2 * 15%</td>
                                <td colspan="2" align="right"><%=formatnumber(tax15_31,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="4" height="30" align="center" style="background:#f8f8f8;">15-2 �������ݾ� > �ſ�ī�����6</th>
                                <td colspan="5" align="left" >6 * 15% + [(15-2) - 6] * 30%</td>
                                <td colspan="2" align="right"><%=formatnumber(tax15_32,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="11" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">16 ���</th>
                            </tr>
                            <tr>
                                <th colspan="3" height="30" align="center" style="background:#f8f8f8;">����</th>
                                <th align="center" style="background:#f8f8f8;">�����Ⱓ</th>
                                <th colspan="2" align="center" style="background:#f8f8f8;">�ݾ�</th>
                                <th colspan="5" align="left" style="background:#f8f8f8;">16üũī�� �� ���� ������ ������</th>
							</tr>
                            <tr>
                                <th rowspan="2" colspan="3" height="30" align="center" style="background:#f8f8f8;">������ �ſ�ī��� ����</th>
                                <th align="center" style="background:#f8f8f8;">2013��</th>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td rowspan="2" colspan="5" align="left" >&nbsp;</td>
							</tr>
                            <tr>
                                <th height="30" align="center" style=" border-left:1px solid #e3e3e3; background:#f8f8f8;">2014��</th>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <th rowspan="2" colspan="3" height="30" align="center" style="background:#f8f8f8;">������ �ſ�ī��� ����</th>
                                <th align="center" style="background:#f8f8f8;">2013��</th>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
                                <td rowspan="2" colspan="5" align="left" >(����)<br>��2013�� ������ �ſ�ī�� �� ���� �� 2014�� ������ �ſ�ī�� �� ���� : "0"<br>��2013�� ������ �ſ�ī�� �� ���� < 2014�� ������ �ſ�ī���<br>����:(2014�� �Ϲݱ� �߰����������� - 2013�� �߰�����������<br>*50%) * 10%(��, ������ ��� "0")</td>
							</tr>
                            <tr>
                                <th height="30" align="center" style="border-left:1px solid #e3e3e3; background:#f8f8f8;">2014��<br>�Ϲݱ�</th>
                                <td colspan="2" align="right"><%=formatnumber(c_hap1,0)%>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="11" height="30" align="center">&nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="11" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">������Ư�����ѹ� ����ɡ���121����2��8�׿� ���� �ſ�ī�� �� ���ݾ׿� ���� �ҵ������ ��û �մϴ�.<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2015 �� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��û�� : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(���� �Ǵ� ��)<br></td>
                            </tr>
                            <tr>
                                <td colspan="11" height="30" align="center">&nbsp;</td>
                            </tr>
                            <tr>
                                <th height="30" align="center" style="background:#f8f8f8;">���񼭷�</th>
                                <td colspan="9" height="30" align="left"  >�ſ�ī�� �� ���ݾ� Ȯ�μ�(���� ��74ȣ��5������ ���մϴ�) �Ǵ� ����û Ȩ���������� �����ϴ� �ſ�ī�� �� ���ݾ� ���� ����� ���� 1��</td>
                                <td>������ ����</td>
							</tr>
                   <%
				          end if
				   %>                        
						</tbody>
					</table>
				</div>
				<table width="1150" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<br>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				    <br>                 
                    </td>
			      </tr>
				</table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

