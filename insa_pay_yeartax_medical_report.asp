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
emp_person = cstr(rs_emp("emp_person1")) + "-" + cstr(rs_emp("emp_person2"))	
rs_emp.close()	


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

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
Rs.Open Sql, Dbconn, 1


title_line = "��������-�Ƿ�����޸�"
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
				<form action="insa_pay_yeartax_medical_report.asp" method="post" name="frm">
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
							  <th class="left" style=" border-top:1px solid #e3e3e3;">�缺��</th>
                              <td><%=emp_name%></td>
                              <th class="left" style=" border-top:1px solid #e3e3e3;">���ֹε�Ϲ�ȣ(�Ǵ� �ܱ��ε�Ϲ�ȣ)</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th class="left">����θ�</th>
                              <td><%=company_name%></td>
                              <th class="left">���ü��</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
                              <td colspan="4">(<%=inc_yyyy%>) �� �Ƿ�� ���޸�</td>
						    </tr>
						</thead>
					</table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="*" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">����ó</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">���޸�</th>
                              </tr>
                              <tr>
                                <th class="first" scope="col">�����ڵ�</th>
                                <th scope="col">����</th>
                                <th scope="col">���ֹε�Ϲ�ȣ</th>
                                <th scope="col">�캻�ε�<br>�ش翩��</th>
                                <th scope="col">�����ڵ�Ϲ�ȣ</th>
                                <th scope="col">���ȣ</th>
                                <th scope="col">���Ƿ������ڵ�</th>
                                <th scope="col">��Ǽ�</th>
                                <th scope="col">��ݾ�</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             m_rel = ""
							 if rs("m_rel") = "����" then 
							          m_rel = "0"
							    elseif rs("m_rel") = "��" or rs("m_rel") = "��" or rs("m_rel") = "����" or rs("m_rel") = "����" then 
							                 m_rel = "1"
									   elseif rs("m_rel") = "����" or rs("m_rel") = "���" then 
							                        m_rel = "2"
											  elseif rs("m_rel") = "����" or rs("m_rel") = "�Ƴ�" then 
							                               m_rel = "3"
													 elseif rs("m_rel") = "�Ƶ�" or rs("m_rel") = "��" then 
							                                      m_rel = "4"
														    elseif rs("m_rel") = "����" or rs("m_rel") = "�ճ�" then 
							                                             m_rel = "5"
																   elseif rs("m_rel") = "��(�����ڸ�)" or rs("m_rel") = "��(�����ڸ�)" or rs("m_rel") = "��(�����ڸ�)" or rs("m_rel") = "��(�����ڸ�)" then 
							                                                    m_rel = "6"
																		  elseif rs("m_witak") = "Y" then
																		               m_rel = "7"
																				 elseif rs("m_pensioner") = "Y" then
																		               m_rel = "8"
							 end if
							 m_bon = ""
							 if rs("m_rel") = "����" or rs("m_disab") = "Y" or rs("m_age65") = "Y" then 	
							          m_bon = "1"
								else  
								      m_bon = "2"
							 end if	
							 m_data_gubun = ""
							 if rs("m_data_gubun") = "����û"	then
							 		  m_data_gubun	= "1"
								elseif rs("m_data_gubun") = "���ΰǰ��������"	then
							 		         m_data_gubun	= "2"
									   elseif rs("m_data_gubun") = "�����/������"	then
							 		                m_data_gubun	= "3"
											  elseif rs("m_data_gubun") = "�����޿�"	then
							 		                       m_data_gubun	= "4"
											         elseif rs("m_data_gubun") = "��Ÿ�Ƿ�񿵼���"	then
							 		                              m_data_gubun	= "5"
																  
							 end if		  						   
	           			%>
							<tr>
                                <td><%=m_rel%>&nbsp;</td>
                                <td><%=rs("m_national")%>&nbsp;</td>
                                <td><%=rs("m_person_no")%>&nbsp;</td>
                                <td><%=m_bon%>&nbsp;</td>
                                <td><%=rs("m_trade_no")%>&nbsp;</td>
                                <td><%=rs("m_trade_name")%>&nbsp;</td>
                                <td><%=m_data_gubun%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("m_cnt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("m_amt"),0)%>&nbsp;</td>
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
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_medical_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_medial_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">�Ƿ�����޸� ���</a>
                    <a href="insa_pay_yeartax_donation_report.asp" class="btnType04">��αݸ���</a>
                    <a href="insa_pay_yeartax_credit_report.asp" class="btnType04">�ſ�ī��� ����</a>
                    <a href="insa_pay_yeartax_tax_report.asp" class="btnType04">�ҵ�����Ű�</a>
                    <a href="insa_pay_yeartax_saving_report.asp" class="btnType04">����.���� �� �ҵ漼�� ��������</a>
                    <a href="insa_pay_yeartax_house_report.asp" class="btnType04">�����������ӱݿ����ݻ�ȯ ����</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

