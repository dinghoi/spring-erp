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

sql = "select * from pay_yeartax_donation where d_year = '"&inc_yyyy&"' and d_emp_no = '"&emp_no&"' ORDER BY d_emp_no,d_person_no,d_seq ASC"
Rs.Open Sql, Dbconn, 1


title_line = "��������-��α� ����"
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
				<form action="insa_pay_yeartax_donation_report.asp" method="post" name="frm">
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
                              <td colspan="4">(<%=inc_yyyy%>) �� ��α� ���޸�</td>
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
                              <col width="12%" >
                              <col width="12%" >
                              <col width="*" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">����ó ��������</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�����ڵ�</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�Ǽ�</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�ݾ�</th>
                              </tr>
                              <tr>
                                <th class="first" scope="col">�����ڵ�</th>
                                <th scope="col">���ܱ���</th>
                                <th scope="col">�ֹε�Ϲ�ȣ</th>
                                <th scope="col">����</th>
                                <th scope="col">����ڵ�Ϲ�ȣ</th>
                                <th scope="col">��ȣ</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             
							 d_year = rs("d_year")
							 d_emp_no = rs("d_emp_no")
							 d_person_no = rs("d_person_no")
							 
							 sql = "select * from pay_yeartax_family where f_year = '"&d_year&"' and f_emp_no = '"&d_emp_no&"' and f_person_no = '"&d_person_no&"'"
                             rs_fami.Open Sql, Dbconn, 1
                             if not rs_fami.eof then
							        f_national = rs_fami("f_national")
							        f_pensioner = rs_fami("f_pensioner")
							        f_witak = rs_fami("f_witak")
							    else
								    f_national = ""
							        f_pensioner = ""
							        f_witak = ""
							 end if
						     rs_fami.close()	
							 
							 d_rel = ""
							 if rs("d_rel") = "����" then 
							          d_rel = "1"
							    elseif rs("d_rel") = "����" or rs("d_rel") = "�Ƴ�" then 
							                 d_rel = "2"
								       elseif rs("d_rel") = "�Ƶ�" or rs("d_rel") = "��" then 
							                        d_rel = "3"
								              elseif rs("d_rel") = "��" or rs("d_rel") = "��" or rs("d_rel") = "����" or rs("d_rel") = "����" then
							                               d_rel = "4"
												     elseif rs("d_rel") = "��(�����ڸ�)" or rs("d_rel") = "��(�����ڸ�)" or rs("d_rel") = "��(�����ڸ�)" or rs("d_rel") = "��(�����ڸ�)" then 
							                                      d_rel = "5"
														    else
															      d_rel = "6"
							 end if
				             d_data_gubun = ""
							 if rs("d_data_gubun") = "������α�"	then
							 		  d_data_gubun	= "10"
								elseif rs("d_data_gubun") = "��ġ�ڱݱ�α�"	then
							 		         d_data_gubun	= "20"
									   elseif rs("d_data_gubun") = "������α�"	then
							 		                d_data_gubun	= "40"
											  elseif rs("d_data_gubun") = "������ü������α�"	then
							 		                       d_data_gubun	= "41"
											         elseif rs("d_data_gubun") = "�츮�������ձ�α�"	then
							 		                              d_data_gubun	= "42"
														    elseif rs("d_data_gubun") = "������ü��������α�"	then
							 		                                     d_data_gubun	= "50"
																  
							 end if	
	           			%>
							<tr>
                                <td><%=d_rel%>&nbsp;</td>
                                <td><%=f_national%>&nbsp;</td>
                                <td><%=rs("d_person_no")%>&nbsp;</td>
                                <td><%=rs("d_name")%>&nbsp;</td>
                                <td><%=rs("d_trade_no")%>&nbsp;</td>
                                <td><%=rs("d_trade_name")%>&nbsp;</td>
                                <td><%=d_data_gubun%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("d_cnt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("d_amt"),0)%>&nbsp;</td>
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
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType04">�Ƿ�����޸���</a>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_donation_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_donation_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">��α� ���� ���</a>
                    <a href="insa_pay_yeartax_credit_report.asp" class="btnType04">�ſ�ī��� ����</a>
                    <a href="insa_pay_yeartax_tax_report.asp" class="btnType04">�ҵ�����Ű�</a>
                    <a href="insa_pay_yeartax_saving_report.asp" class="btnType04">����.���� �� �ҵ漼�� ��������</a>
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

