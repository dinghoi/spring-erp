<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_medical.asp"

y_final=Request("y_final")

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

for i = 1 to 10
    family_tab(i,1) = ""
	family_tab(i,2) = ""
	family_tab(i,3) = ""
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
emp_person = cstr(rs_emp("emp_person1")) + cstr(rs_emp("emp_person2"))	
rs_emp.close()	

tot_pay = 0
y_nhis_amt = 0
y_longcare_amt = 0
y_epi_amt = 0
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
'Set rs_year = DbConn.Execute(SQL)
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
y_nhis_amt = y_nhis_amt + y_longcare_amt
y_nhis_tax = y_nhis_amt
y_epi_tax = y_epi_amt

b_nhis = 0
b_longcare = 0
b_epi = 0
Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
'Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
       b_nhis = b_nhis + rs_bef("b_nhis")
	   b_longcare = b_longcare + rs_bef("b_longcare")
	   b_epi = b_epi + rs_bef("b_epi")
	   tot_pay = tot_pay + rs_bef("b_pay") + rs_bef("b_bonus") + rs_bef("b_deem_bonus")
	rs_bef.MoveNext()
loop
rs_bef.close()
b_nhis = b_nhis + b_longcare
b_nhis_tax = b_nhis
b_epi_tax = b_epi

tot_3per = int(tot_pay * (3 / 100))

bon65_tot = 0
disab_tot = 0
other_tot = 0
Sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
rs_medi.Open Sql, Dbconn, 1
'Set rs_medi = DbConn.Execute(SQL)
do until rs_medi.eof
       if rs_medi("m_disab") = "Y" then
	         disab_tot = disab_tot + rs_medi("m_amt")
		  elseif rs_medi("m_age65") = "Y" or rs_medi("m_rel") = "����" then
		             bon65_tot = bon65_tot + rs_medi("m_amt")
				 else
	                 other_tot = other_tot + rs_medi("m_amt")
		end if
	rs_medi.MoveNext()
loop
rs_medi.close()

tot_amt = bon65_tot + disab_tot + other_tot 
bon65_tax = bon65_tot
disab_tax = disab_tot
other_tax = other_tot
tot_tax = bon65_tax + disab_tax + other_tax 

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
'Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "����" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
		  i = i + 1
		  family_tab(i,1) = rs_fami("f_rel")
	      family_tab(i,2) = rs_fami("f_family_name")
	      family_tab(i,3) = rs_fami("f_person_no")
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�������� - Ư������(�Ƿ��) "
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
				<form action="insa_pay_yeartax_medical.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="16%" >
							<col width="14%" >
							<col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (�Ի���:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th style=" border-bottom:1px solid #e3e3e3;">�Ҽ�(<%=emp_company%><input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true">)</th>
							  <td colspan="4" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_org_name%>
                                <input name="emp_org_name" type="hidden" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                - <%=emp_grade%>
                                <input name="emp_grade" type="hidden" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                - <%=emp_position%>
                                <input name="emp_position" type="hidden" value="<%=emp_position%>" style="width:70px" readonly="true">
                                (�ͼӳ⵵:
                                <input name="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:40px; text-align:center" readonly="true">)
                                &nbsp;&nbsp;&nbsp;&nbsp;
                                �ѱ޿�:&nbsp;<%=formatnumber(tot_pay,0)%>���� 3%�ݾ���&nbsp;<%=formatnumber(tot_3per,0)%>���Դϴ�.
                              </td>
						    </tr>
                             <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���ⱸ��</th>
                              <th>�ݾ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ѵ���</th>
                              <th>������</th>
						    </tr>
                            <tr>
							  <th rowspan="4">�Ƿ��</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">���Ρ�65���̻���</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(bon65_tot,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ��������</th>
                              <td class="right"><%=formatnumber(bon65_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����� �Ƿ��</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(disab_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�ۼ��������</th>
                              <td class="right"><%=formatnumber(disab_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�� ���� ��������� �Ƿ��</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�����</th>
                              <td class="right"><%=formatnumber(other_tot,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�ۼ��������</th>
                              <td class="right"><%=formatnumber(other_tax,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;">�Ƿ�� ��</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(tot_tax,0)%>&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">�� ���踦 �����ϴ� �ξ簡��(����.�ҵ�ݾ׿�ǿ� ���Ѿ���)�� ������ �Ƿ��μ� �ѱ޿����� 3%�� �ʰ��� �Ƿ�� ���ؼ��� ��������.<br>
                �� 65���̻��� ��.������� ���Ͽ� ������ �Ƿ��� �Ƿ�����޾��� �������Դϴ�.<br>
                �� �̿뼺�������� �� �ǰ������Ǿ�ǰ ���Ժ�� �Ƿ������׿��� �����մϴ�.<br>
                �� ���� �ѱ޿����� 3%�ݾ��� Ȯ�� �Ͻð� �Ƿ�� ���հ� �ݾ��� 3%�̸��̸� �Է��� �ʿ� ����.</h3>

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="5%" >
                              <col width="5%" >
                              <col width="12%" >
                              <col width="10%" >
                              <col width="*" >
                              <col width="4%" >
                              <col width="10%" >
                              <col width="5%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">����</th>
                                <th scope="col">����(����)</th>
                                <th scope="col">������̸�</th>
                                <th scope="col">�ֹε�Ϲ�ȣ</th>
                                <th scope="col">���<br>����</th>
                                <th scope="col">65��<br>�̻���</th>
                                <th scope="col">�Ƿ�������ڵ�</th>
                                <th scope="col">����ڵ�Ϲ�ȣ</th>
                                <th scope="col">��ȣ��</th>
                                <th scope="col">�Ǽ�</th>
                                <th scope="col">�ݾ�</th>
                                <th scope="col">�Ȱ��<br>���Կ���</th>
                                <th scope="col">���</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             m_disab = rs("m_disab")
							 m_age65 = rs("m_age65")
							 m_eye = rs("m_eye")
	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("m_rel")%>&nbsp;</td>
                                <td><%=rs("m_name")%>&nbsp;</td>
                                <td><%=mid(cstr(rs("m_person_no")),1,6)%>-<%=mid(cstr(rs("m_person_no")),7,7)%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="m_disab" value="Y" <% if m_disab = "Y" then %>checked<% end if %> id="m_disab"></td>
                                <td>
                                <input type="checkbox" name="m_age65" value="Y" <% if m_age65 = "Y" then %>checked<% end if %> id="m_age65"></td>
                                <td><%=rs("m_data_gubun")%>&nbsp;</td>
                                <td><%=rs("m_trade_no")%>&nbsp;</td>
                                <td><%=rs("m_trade_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("m_cnt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("m_amt"),0)%>&nbsp;</td>
                                <td>
                                <input type="checkbox" name="m_eye" value="Y" <% if m_eye = "Y" then %>checked<% end if %> id="m_eye"></td>
                        <% if y_final <> "Y" then  %>                                  
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_medical_add.asp?m_year=<%=rs("m_year")%>&m_emp_no=<%=rs("m_emp_no")%>&m_seq=<%=rs("m_seq")%>&m_person_no=<%=rs("m_person_no")%>&m_emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_pay_yeartax_medical_add_pop','scrollbars=yes,width=850,height=370')">����</a></td>
                        <%    else  %>
                                <td>&nbsp;</td>
                        <% end if  %>										
                            </tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_insurance.asp" class="btnType04">�������</a>
              <% if y_final <> "Y" then  %>                     
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_medical_add.asp?m_year=<%=inc_yyyy%>&m_emp_no=<%=emp_no%>&m_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_medical_add_pop','scrollbars=yes,width=850,height=370')" class="btnType04">�Ƿ���߰����</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_medical.asp" class="btnType04">�Ƿ����</a>
			  <%   end if  %>                          
                    <a href="insa_pay_yeartax_edu.asp" class="btnType04">��������</a>
                    <a href="insa_pay_yeartax_house.asp" class="btnType04">�����ڱݵ��</a>
                    <a href="insa_pay_yeartax_donation.asp" class="btnType04">��αݵ��</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">                
			</form>
		</div>				
	</div>        				
	</body>
</html>

