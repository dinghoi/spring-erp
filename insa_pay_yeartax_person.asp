<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim year_tab(3,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_person.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	inc_yyyy = request.form("inc_yyyy")
  else
	inc_yyyy = request("inc_yyyy")
end if

if view_condi = "" then
	'inc_yyyy = mid(cstr(now()),1,4)
	inc_yyyy = cint(mid(now(),1,4)) - 1
	ck_sw = "n"
end if

' �ֱ�3���⵵ ���̺�� ����
'year_tab(3,1) = mid(now(),1,4)
'year_tab(3,2) = cstr(year_tab(3,1)) + "��"
'year_tab(2,1) = cint(mid(now(),1,4)) - 1
'year_tab(2,2) = cstr(year_tab(2,1)) + "��"
'year_tab(1,1) = cint(mid(now(),1,4)) - 2
'year_tab(1,2) = cstr(year_tab(1,1)) + "��"

' �ֱ�3���⵵ ���̺�� ����
year_tab(3,1) = cint(mid(now(),1,4)) - 1
year_tab(3,2) = cstr(year_tab(3,1)) + "��"
year_tab(2,1) = cint(mid(now(),1,4)) - 2
year_tab(2,2) = cstr(year_tab(2,1)) + "��"
year_tab(1,1) = cint(mid(now(),1,4)) - 3
year_tab(1,2) = cstr(year_tab(1,1)) + "��"

Set Dbconn=Server.CreateObject("ADODB.Connection")
DBConn.open DbConnect

Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
'Set rs_emp = Server.CreateObject("ADODB.Recordset")
'Set rs_year = Server.CreateObject("ADODB.Recordset")
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



Sql = "select emp_in_date, emp_name, emp_grade, emp_position, emp_company, emp_org_name from emp_master where emp_no = '"&emp_no&"'"
Set rs_emp = DBConn.Execute(Sql)

emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")

rs_emp.close()

sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"

Set rs_year = DBConn.Execute(Sql)

If rs_year.EOF Or rs_year.BOF Then
	y_basic_cnt = ""
	bon_person_no = ""
	y_woman = ""
	y_single = ""
Else
	y_basic_cnt =  rs_year("y_wife") + rs_year("y_age20_cnt") + rs_year("y_age60_cnt")
	bon_person_no =  cstr(rs_year("y_person_no1")) & cstr(rs_year("y_person_no2"))

	if rs_year("y_woman") = "Y" then
		y_woman = "��"
	else
		y_woman = ""
	end If

	if rs_year("y_single") = "Y" then
		y_single = "��"
	else
		y_single = ""
	end If

	y_nhis_amt = rs_year("y_nhis_amt")
	y_epi_amt = rs_year("y_epi_amt")
	y_longcare_amt = rs_year("y_longcare_amt")
	y_support_cnt = rs_year("y_support_cnt")
	y_daja_cnt = rs_year("y_daja_cnt")
	y_old_cnt = rs_year("y_old_cnt")
	y_holt_cnt = rs_year("y_holt_cnt")
	y_disab_cnt = rs_year("y_disab_cnt")
	y_age6_cnt = rs_year("y_age6_cnt")
	y_emp_no = rs_year("y_emp_no")

End If

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

'sum_nts_insuran = sum_nts_insuran + rs_year("y_nhis_amt") + rs_year("y_epi_amt") + rs_year("y_longcare_amt")
sum_nts_insuran = sum_nts_insuran + y_nhis_amt + y_epi_amt + y_longcare_amt


sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�������� - �������� �� �ҵ������ "

d_chk = "1"
b_bonus = 23045000
b_pay = 1230000

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
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}

			function form_chk(){
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_person.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <label>
                             <strong>��� : </strong>
                                <input name="emp_no" type="text" value="<%=emp_no%>" style="width:50px" readonly="true">
                                -
                                <input name="emp_name" type="text" value="<%=emp_name%>" style="width:60px" readonly="true">
                                </label>
                                <label>
                             <strong>���� : </strong>
                                <input name="emp_grade" type="text" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                -
                                <input name="emp_position" type="text" value="<%=emp_position%>" style="width:70px" readonly="true">
                                </label>
                                <label>
                             <strong>�Ի��� : </strong>
                                <input name="emp_in_date" type="text" value="<%=emp_in_date%>" style="width:70px" readonly="true">
                                </label>
                                <label>
                             <strong>�Ҽ� : </strong>
                                <input name="emp_company" type="text" value="<%=emp_company%>" style="width:90px" readonly="true">
                                -
                                <input name="emp_org_name" type="text" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                </label>
                             <strong>�ͼӳ⵵ : </strong>
                                <select name="inc_yyyy" id="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:70px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If inc_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="10%" >
						<col width="15%" >
						<col width="10%" >
						<col width="15%" >
						<col width="12%" >
						<col width="13%" >
                        <col width="12%" >
						<col width="13%" >
					</colgroup>
				    <tbody>
                    <tr>
                      <td colspan="8" class="left">�� ������������ �ֹε�ϵ(����), ������������(�����), ����� �������� �ų� �����Ͽ��� �մϴ�.<br>�� �����, �Ƿ��, ������, �ſ�ī��� ����, ��α��� ����Ͻø� �� ȭ�鿡�� ��ȸ�� �� �ֽ��ϴ�.<br>�� �������� Ȯ�� : ������ �ʿ��� ���� �λ���� �������׿��� �߰� �� ���� �Ͻñ� �ٶ��ϴ�.<br>�� ����� ������</td>
                    </tr>
                   	<tr>
                      <th colspan="2" style=" border-left:1px solid #e3e3e3;">����</th>
                      <th colspan="4">���⼭��</th>
                      <th colspan="2">�߱�ó</th>
                    </tr>
                    <tr>
                      <td colspan="2" class="center">�����</td>
                      <td colspan="4" class="left">����� ����, ����ε����(��ø)�纻, ����ī��</td>
                      <td colspan="2" class="left">��������(www.egov.go.kr) �Ǵ� ���鵿 �ֹμ�Ÿ</td>
                    </tr>
                    <tr>
                      <td colspan="2" class="center">������</td>
                      <td colspan="4" class="left">�������� �纻, ����û ��ø</td>
                      <td colspan="2" class="left">����û</td>
                    </tr>
                    <tr>
                      <td colspan="2" class="center">�׽� ġ�Ḧ ���ϴ� ����ȯ��</td>
                      <td colspan="4" class="left">�������� �����ϴ� ����� ����(�ǻ��� ���� �Ǵ� ���� �Ͽ��� ��)</td>
                      <td colspan="2" class="left">�ش� �Ƿ���</td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left" style=" border-left:1px solid #ffffff;">&nbsp;</td>
                    </tr>
			        </tbody>
			      </table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="10%" >
							<col width="2%" >
                            <col width="2%" >
							<col width="3%" >
                            <col width="3%" >
                            <col width="7%" >

                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="7%" >
                            <col width="7%" >
						</colgroup>
						<thead>
                            <tr>
								<th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3;">���������׸�</th>
                                <th colspan="10" scope="col" style=" border-bottom:1px solid #e3e3e3;">���� �ҵ���� �׸�</th>
							</tr>
                            <tr>
								<th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�⺻<br>����</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">���<br>���</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">���<br>�Ծ�</th>
                                <th rowspan="2" scope="col">����</th>
                                <th rowspan="2" scope="col">�����<br>(�ǰ������� ����)</th>
                                <th rowspan="2" scope="col">�Ƿ��</th>
                                <th rowspan="2" scope="col">������</th>
                                <th rowspan="2" scope="col">�ſ�ī��<br>(������塤���߱�������)</th>
                                <th rowspan="2" scope="col">����ī�� ��<br>(������塤���߱�������)</th>
                                <th rowspan="2" scope="col">���ݿ�����<br>(������塤���߱�������)</th>
                                <th rowspan="2" scope="col">�������<br>����</th>
                                <th rowspan="2" scope="col">���߱���<br>�̿��</th>
                                <th rowspan="2" scope="col">��α�</th>
							</tr>
                            <tr>
								<th class="first" scope="col">����</th>
                                <th scope="col">�ֹε�Ϲ�ȣ</th>
								<th scope="col">�γ���</th>
                                <th scope="col">�Ѻθ�</th>
                                <th scope="col">�����</th>
                                <th scope="col">6������</th>
							</tr>
						</thead>
						<tbody>
                            <tr>
                                <td rowspan="2" colspan="2" style=" border-top:2px solid #515254;">�ο� :&nbsp;<%=y_support_cnt%><%'=rs_year("y_support_cnt")%>&nbsp;(���ڳ�:&nbsp;<%'=rs_year("y_daja_cnt")%><%=y_daja_cnt%>��)</td>
                                <td colspan="2" style=" border-top:2px solid #515254;"><%=y_basic_cnt%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=y_old_cnt%><%'=rs_year("y_old_cnt")%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;"><%=y_holt_cnt%><%'=rs_year("y_holt_cnt")%>&nbsp;</td>
                                <td style=" border-top:2px solid #515254;">����û�ڷ�</td>
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
                                <td><%=y_disab_cnt%><%'=rs_year("y_disab_cnt")%>&nbsp;</td>
                                <td><%=y_age6_cnt%><%'=rs_year("y_age6_cnt")%>&nbsp;</td>
                                <td>�׹����ڷ�</td>
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
							       hap_nts_insuran = rs("i_ilban_nts") + rs("i_disab_nts") + rs_year("y_nhis_amt") + rs_year("y_epi_amt") + rs_year("y_longcare_amt")
								   hap_oth_insuran = rs("i_ilban_other") + rs("i_disab_other")
							   else
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
				</div>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">
			</form>
		</div>
	</div>
	</body>
</html>

