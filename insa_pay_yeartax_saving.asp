<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_saving.asp"

y_final=Request("y_final")
s_id=Request("s_id")

'response.write(s_id)

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

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ann = Server.CreateObject("ADODB.Recordset")
Set rs_savi = Server.CreateObject("ADODB.Recordset")
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
rs_emp.close()	

y_nps_other = 0

y_nps_amt = 0
Sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
Set rs_year = DbConn.Execute(SQL)
if not rs_year.eof then
       y_nps_amt = rs_year("y_nps_amt")
   else
       y_nps_amt = 0
end if
y_nps_tax = y_nps_amt

b_nps = 0
Sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
rs_bef.Open Sql, Dbconn, 1
Set rs_bef = DbConn.Execute(SQL)
do until rs_bef.eof
       b_nps = b_nps + rs_bef("b_nps")
	rs_bef.MoveNext()
loop
rs_bef.close()
b_nps_tax = b_nps

if s_id = "��������" then
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
end if

if s_id = "���ø�������" then
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

      oj_tot_amt = tot_cheng + tot_jutak + tot_gunro + tot_jangi
      oj_tot_tax = tax_cheng + tax_jutak + tax_gunro + tax_jangi
end if

if s_id = "����ֽ�������" then
      tot_2 = 0
      tot_3 = 0
      tot_4 = 0
      Sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
      rs_savi.Open Sql, Dbconn, 1
      Set rs_savi = DbConn.Execute(SQL)
      do until rs_savi.eof
            if rs_savi("s_type") = "2����" then 
	                 tot_2 = tot_2 + rs_savi("s_amt")
		       elseif rs_savi("s_type") = "3����" then 
	                        tot_3 = tot_3 + rs_savi("s_amt")
			          elseif rs_savi("s_type") = "4����" then 
	                              tot_4 = tot_4 + rs_savi("s_amt")
		    end if
	        rs_savi.MoveNext()
      loop
      rs_savi.close()

      tax_2 = tot_2
      tax_3 = tot_3
      tax_4 = tot_4

      ojj_tot_amt = tot_2 + tot_3 + tot_4 
      ojj_tot_tax = tax_2 + tax_3 + tax_4 
end if



sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�������� - �׹��ǰ���(" + s_id + ")"
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
				<form action="insa_pay_yeartax_saving.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="14%" >
							<col width="14%" >
							<col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td colspan="2" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (�Ի���:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th style=" border-bottom:1px solid #e3e3e3;">�Ҽ�(<%=emp_company%><input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true">)</th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_org_name%>
                                <input name="emp_org_name" type="hidden" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                - <%=emp_grade%>
                                <input name="emp_grade" type="hidden" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                - <%=emp_position%>
                                <input name="emp_position" type="hidden" value="<%=emp_position%>" style="width:70px" readonly="true">
                                (�ͼӳ⵵:
                                <input name="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:40px; text-align:center" readonly="true">)
                              </td>
                            </tr>
                     <% if s_id = "��������" then  %>
                             <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���ⱸ��</th>
                              <th>�ݾ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ѵ���</th>
                              <th>������</th>
						    </tr>
                            <tr>
							  <th rowspan="4"><%=s_id%></th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">���ο�������(2000������)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_2000,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Ծ���40%�� 72����</th>
                              <td class="right"><%=formatnumber(tax_2000,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��������(2001������)</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_2001,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_2001,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������ݼҵ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_endi,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_endi,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;"><%=s_id%> ��</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(oy_tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(oy_tot_tax,0)%>&nbsp;</td>
						    </tr>
                     <% end if %>	
                     <% if s_id = "���ø�������" then  %> 
                             <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���ⱸ��</th>
                              <th>�ݾ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ѵ���</th>
                              <th>������</th>
						    </tr>
                            <tr>
							  <th rowspan="5"><%=s_id%></th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">û������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_cheng,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_cheng,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ٷ������ø�������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_gunro,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_gunro,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">����û����������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_jutak,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_jutak,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">������ø�������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(tot_jangi,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(tax_jangi,0)%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3;"><%=s_id%> ��</th>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(oj_tot_amt,0)%>&nbsp;</td>
                              <th style="background:#f8f8f8;">&nbsp;</th>
                              <td class="right"><%=formatnumber(oj_tot_tax,0)%>&nbsp;</td>
						    </tr>                     
                     <% end if %>	
                     <% if s_id = "����ֽ�������" then  %> 
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���ⱸ��</th>
                              <th>�ݾ�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ѵ���</th>
                              <th>������</th>
						    </tr>
                            <tr>
							  <th ><%=s_id%></th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">����ֽ�������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><%=formatnumber(ojj_tot_amt,0)%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">�ۼ���� ����</th>
                              <td class="right"><%=formatnumber(ojj_tot_tax,0)%>&nbsp;</td>
						    </tr>
                     <% end if %>	              	
						</thead>
						<tbody>
					</table>
          <% if s_id = "��������" then  %>                    
				<h3 class="stit">�� ���ο�������� ���������� ��� �����Ѱ��, ���� ������ �����Ͽ� �ҵ������ ���� �� ����<br>
                �� �������� �����ѵ��� 400����<br>
                �� ���ο���/���������� �ٷ��� ���θ��Ƿ� ������ ��쿡�� �������</h3>
          <% end if %>	  
          <% if s_id = "���ø�������" then  %>  
                <h3 class="stit">�� ���������� ��� ������ �������� ���� �����ַ� ���θ��Ƿ� ������� ���ø��� ���࿡ ������ �ٷ��ڸ� �������.<br>
                �� ����� �Ǵ� �ξ簡���� ���� �ܵ� �����ֵ� ��������<br>
                �� ���ø��������� ��� �ݵ�� �����ֿ��� ��������.</h3>
          <% end if %>	
          <% if s_id = "����ֽ�������" then  %> 
                <h3 class="stit">�� �����ѵ�: ���Կ����� 1,200����(�б⺰ 300�����ѵ�).<br>
                �� �����ݾ� : 2���� ���Աݾ��� 10%, 3���� ���Աݾ��� 5% </h3> 
          <% end if %>	           

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="20%" >
                              <col width="16%" >
                              <col width="20%" >
                              <col width="20%" >
                              <col width="16%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">����</th>
                                <th scope="col">����</th>
                                <th scope="col">�������</th>
                                <th scope="col">�������</th>
                                <th scope="col">����/���ǹ�ȣ</th>
                                <th scope="col">�ݾ�</th>
                                <th scope="col">���</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                  
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_saving_add.asp?s_year=<%=rs("s_year")%>&s_emp_no=<%=rs("s_emp_no")%>&s_seq=<%=rs("s_seq")%>&s_emp_name=<%=rs("s_emp_name")%>&s_id=<%=s_id%>&u_type=<%="U"%>','insa_pay_yeartax_saving_add_pop','scrollbars=yes,width=750,height=300')">����</a></td>
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
              <% if s_id = "��������" then  %>
					<div class="btnRight">
              <% if y_final <> "Y" then  %>     					
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_saving_add.asp?s_year=<%=inc_yyyy%>&s_emp_no=<%=emp_no%>&s_emp_name=<%=emp_name%>&s_id=<%=s_id%>&u_type=<%=""%>','insa_pay_yeartax_saving_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">�������� �����׸��Է�</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="��������"%>" class="btnType04">��������</a>
			  <%   end if  %>                     
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="���ø�������"%>" class="btnType04">���ø�������</a>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="����ֽ�������"%>" class="btnType04">����ֽ�������</a>
					</div> 
			  <% end if %>		
              <% if s_id = "���ø�������" then  %>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="��������"%>" class="btnType04">��������</a>
              <% if y_final <> "Y" then  %>                         
					<a href="#" onClick="pop_Window('insa_pay_yeartax_saving_add.asp?s_year=<%=inc_yyyy%>&s_emp_no=<%=emp_no%>&s_emp_name=<%=emp_name%>&s_id=<%=s_id%>&u_type=<%=""%>','insa_pay_yeartax_saving_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">���ø������� �����׸��Է�</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="���ø�������"%>" class="btnType04">���ø�������</a>
			  <%   end if  %>                        
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="����ֽ�������"%>" class="btnType04">����ֽ�������</a>
					</div> 
			  <% end if %>	
              <% if s_id = "����ֽ�������" then  %>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="��������"%>" class="btnType04">��������</a>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="���ø�������"%>" class="btnType04">���ø�������</a>
              <% if y_final <> "Y" then  %>                       
					<a href="#" onClick="pop_Window('insa_pay_yeartax_saving_add.asp?s_year=<%=inc_yyyy%>&s_emp_no=<%=emp_no%>&s_emp_name=<%=emp_name%>&s_id=<%=s_id%>&u_type=<%=""%>','insa_pay_yeartax_saving_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">����ֽ������� �����׸��Է�</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="����ֽ�������"%>" class="btnType04">����ֽ�������</a>
			  <%   end if  %>     					
                    </div> 
			  <% end if %>		                 
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">
				<input type="hidden" name="s_id" value="<%=s_id%>" ID="Hidden1">                
			</form>
		</div>				
	</div>        				
	</body>
</html>

