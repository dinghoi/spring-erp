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

s_id = "��������"

sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
Rs.Open Sql, Dbconn, 1


title_line = "��������-�����ס����ְ� �� �����Ӵ����ӱ� ������ ��ȯ�� �ҵ���� ����"
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
				<form action="insa_pay_yeartax_house_report.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
                            <col width="10%" >
							<col width="30%" >
							<col width="20%" >
							<col width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td rowspan="4">1. ���� ����</td>
                              <th class="left">����θ�</th>
                              <td><%=company_name%></td>
                              <th class="left">���ü��</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�鼺��</th>
                              <td><%=emp_name%></td>
                              <th class="left" style=" border-top:1px solid #e3e3e3;">���ֹε�Ϲ�ȣ(�Ǵ� �ܱ��ε�Ϲ�ȣ)</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ּ�</th>
                              <td colspan="3" class="left"><%=addr_name%><br>(��ȭ��ȣ:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
                            </tr>
                            <tr>
                              <th class="left" style="border-left:1px solid #e3e3e3;">������ ������</th>
                              <td colspan="3" class="left"><%=addr_name%><br>(��ȭ��ȣ:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
						    </tr>
						</thead>
					</table>

                    <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="12%" >
							  <col width="8%" >
							  <col width="8%" >
							  <col width="*" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="12%" >
                              <col width="12%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="9" class="left" scope="col" style="border-bottom:1px solid #e3e3e3;">2. ������ �ҵ���� ��</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col">�Ӵ��μ���<br>(��ȣ)</th>
                                <th rowspan="2" scope="col">�ֹε�Ϲ�ȣ<br>(����ڹ�ȣ)</th>
                                <th rowspan="2" scope="col">��������</th>
                                <th rowspan="2" scope="col">���ð��<br>����(��)</th>
                                <th rowspan="2"scope="col" >�Ӵ�����༭ �� �ּ���</th>
                                <th colspan="2" scope="col" style="border-bottom:1px solid #e3e3e3;">��༭��<br>�Ӵ��� ���Ⱓ</th>
                                <th rowspan="2" scope="col">���� ������(��)</th>
                                <th rowspan="2" scope="col">�����ݾ�(��)</th>
                              </tr>
                              <tr>
                                <th scope="col" style="border-left:1px solid #e3e3e3;">������</th>
                                <th scope="col">������</th>
                              </tr>
                            </thead>
                            <tbody>
                            
						<%
						do until rs.eof
                             if rs("s_type") = "�������ݼҵ����" then
	           			%>
							<tr>
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="left"><%=rs("s_account_no")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
						    end if
							rs.movenext()
						loop
						rs.close()
						%>  
                            <tr>
                                <td colspan="9" class="left" scope="col" style="border-bottom:1px solid #e3e3e3;">�� �������� �����ڵ� - �ܵ�����:1, �ٰ���:2, �ټ�������:3, ��������:4, ����Ʈ:5, ���ǽ���:6, ��Ÿ:7<br><br>
                                �� ��༭�� �Ӵ������Ⱓ - �Խ��ϰ� �������� ���ÿ� ���� ����(����) 2013.01.01.</td>
                            </tr>
						</tbody>
                </table>      

                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="12%" >
							  <col width="*" >
							  <col width="8%" >
                              <col width="12%" >
                              <col width="12%" >
                              <col width="10%" >
                              <col width="14%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="8" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">3 ������ �� �����������Ա� ������ ��ȯ�� �ҵ���� ��</td>
                              </tr>
                              <tr>
                                <td colspan="8" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1) �����Һ���� ��೻��</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col">����</th>
                                <th rowspan="2" scope="col">�ֹε�Ϲ�ȣ</th>
                                <th rowspan="2" scope="col">�����Һ����<br>���Ⱓ</th>
                                <th rowspan="2" scope="col">���Ա�<br>������</th>
                                <th colspan="3" scope="col" style="border-bottom:1px solid #e3e3e3;">������ ��ȯ��</th>
                                <th rowspan="2" scope="col">�����ݾ�(��)</th>
                              </tr>
                              <tr>
                                <th scope="col" style="border-left:1px solid #e3e3e3;">��</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						s_id = "��������"
						sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
                        Rs.Open Sql, Dbconn, 1
						
						do until rs.eof
                             if rs("s_type") <> "�������ݼҵ����" then
	           			%>
							<tr>
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td class="left"><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							end if
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
                </table>   
                
                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="8%" >
							  <col width="*" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="14%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="8" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">2) �Ӵ��� ��೻��</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col">�Ӵ��μ���<br>(��ȣ)</th>
                                <th rowspan="2" scope="col">�ֹε�Ϲ�ȣ<br>(����ڹ�ȣ)</th>
                                <th rowspan="2" scope="col">��������</th>
                                <th rowspan="2" scope="col">���ð��<br>����(��)</th>
                                <th rowspan="2"scope="col" >�Ӵ�����༭ �� �ּ���</th>
                                <th colspan="2" scope="col" style="border-bottom:1px solid #e3e3e3;">��༭��<br>�Ӵ��� ���Ⱓ</th>
                                <th rowspan="2" scope="col">����������(��)</th>
                              </tr>
                              <tr>
                                <th scope="col" style="border-left:1px solid #e3e3e3;">������</th>
                                <th scope="col">������</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						s_id = "���ø�������"
						sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
                        Rs.Open Sql, Dbconn, 1
						
						do until rs.eof

	           			%>
							<tr>
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td><%=rs("s_bank_code")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                            <tr>
                                <td colspan="8" class="left" scope="col" style="border-bottom:1px solid #e3e3e3;">�� �������� �����ڵ� - �ܵ�����:1, �ٰ���:2, �ټ�������:3, ��������:4, ����Ʈ:5, ���ǽ���:6, ��Ÿ:7<br><br>
                                �� ��༭�� �Ӵ������Ⱓ - �Խ��ϰ� �������� ���ÿ� ���� ����(����) 2013.01.01.</td>
                            </tr>
						</tbody>
                   </table>  

                   <table cellpadding="0" cellspacing="0" class="tableList">
                        <colgroup>
							   <col width="100%" >
                        </colgroup>
                        <thead>
                            <tr>
								<td scope="col" style=" border-bottom:2px solid #515254;">�� �� �� ��</td>
							</tr>
                            <tr>
								<td scope="col" class="left" >
                                1. ������ ������ ������ �� ���������ڱ� ���ӱ� ������ ��ȯ�� ������ �޴� �ٷμҵ��ڿ� ���ؼ��� �ش� �ҵ������ ���� ���� �ۼ��ؾ� �մϴ�.<br><br>
                                2. �ش� �Ӵ��� ��ະ�� ���� �հ��� �����ס������ݻ�ȯ�װ� �����ݾ��� ������, �����ݾ��� 0�ΰ�쿡�� ���� �ʽ��ϴ�.<br><br>
                                3. ���������� �ܵ�����, �ٰ���, �ټ�������, ��������, ����Ʈ, ���ǽ���, ��Ÿ �߿��� �ش�Ǵ� �������� �����ڵ带 �����ϴ�.<br><br>
                                4. ������������ �����Ⱓ ������(12.31.) ������ ������������ �����ϴ�.</td>
                            </tr>
                        </thead>
                </table>                        
                     
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType04">�Ƿ�����޸���</a>
                    <a href="insa_pay_yeartax_donation_report.asp" class="btnType04">��αݸ���</a>
                    <a href="insa_pay_yeartax_credit_report.asp" class="btnType04">�ſ�ī��� ����</a>
                    <a href="insa_pay_yeartax_tax_report.asp" class="btnType04">�ҵ�����Ű�</a>
                    <a href="insa_pay_yeartax_saving_report.asp" class="btnType04">����.���� �� �ҵ漼�� ��������</a>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_house_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_donation_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">�����������ӱݿ����ݻ�ȯ ���� ���</a>
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

