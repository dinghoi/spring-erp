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
Set rs_dona = Server.CreateObject("ADODB.Recordset")
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
				<form action="insa_pay_yeartax_house_print.asp" method="post" name="frm">
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
							<col height="30px" width="10%" >
                            <col height="30px" width="10%" >
							<col height="30px" width="30%" >
							<col height="30px" width="20%" >
							<col height="30px" width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td rowspan="4" height="30" align="left">1. ���� ����</td>
                              <th height="30" align="left">����θ�</th>
                              <td height="30" align="center"><%=company_name%></td>
                              <th height="30" align="left">���ü��</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th height="30" align="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�鼺��</th>
                              <td height="30" align="center"><%=emp_name%></td>
                              <th height="30" align="left" style=" border-top:1px solid #e3e3e3;">���ֹε�Ϲ�ȣ(�Ǵ� �ܱ��ε�Ϲ�ȣ)</th>
                              <td height="30" align="center"><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th height="30" align="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���ּ�</th>
                              <td colspan="3" height="30" align="left"><%=addr_name%><br>(��ȭ��ȣ:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
                            </tr>
                            <tr>
                              <th height="30" align="left" style="border-left:1px solid #e3e3e3;">������ ������</th>
                              <td colspan="3" height="30" align="left"><%=addr_name%><br>(��ȭ��ȣ:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
						    </tr>
						</thead>
				  </table>
					<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
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
                                <td colspan="9" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">2. ������ �ҵ���� ��</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">�Ӵ��μ���<br>(��ȣ)</th>
                                <th rowspan="2" scope="col" height="30" align="center">�ֹε�Ϲ�ȣ<br>(����ڹ�ȣ)</th>
                                <th rowspan="2" scope="col" height="30" align="center">��������</th>
                                <th rowspan="2" scope="col" height="30" align="center">���ð��<br>����(��)</th>
                                <th rowspan="2"scope="col" height="30" align="center">�Ӵ�����༭ �� �ּ���</th>
                                <th colspan="2" scope="col" height="30" align="center" style="border-bottom:1px solid #e3e3e3;">��༭��<br>�Ӵ��� ���Ⱓ</th>
                                <th rowspan="2" scope="col" height="30" align="center">���� ������(��)</th>
                                <th rowspan="2" scope="col" height="30" align="center">�����ݾ�(��)</th>
                              </tr>
                              <tr>
                                <th scope="col" height="30" align="center" style="border-left:1px solid #e3e3e3;">������</th>
                                <th scope="col" height="30" align="center">������</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof
                             
							 if rs("s_type") = "�������ݼҵ����" then
	           		 %>
							<tr>
                                <td height="30" align="center"><%=rs("s_type")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_name")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td align="left"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
					<%
							end if
							rs.movenext()
						loop
						rs.close()
						
					%>
                            <tr>
                                <td colspan="9" height="30" align="left" scope="col" style="border-bottom:1px solid #e3e3e3;">�� �������� �����ڵ� - �ܵ�����:1, �ٰ���:2, �ټ�������:3, ��������:4, ����Ʈ:5, ���ǽ���:6, ��Ÿ:7<br><br>
                                �� ��༭�� �Ӵ������Ⱓ - �Խ��ϰ� �������� ���ÿ� ���� ����(����) 2013.01.01.</td>
                            </tr>
						</tbody>
					</table>
                    
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
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
                                <td colspan="8" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">3 ������ �� �����������Ա� ������ ��ȯ�� �ҵ���� ��</td>
                              </tr>
                              <tr>
                                <td colspan="8" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1) �����Һ���� ��೻��</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">����</th>
                                <th rowspan="2" scope="col" height="30" align="center">�ֹε�Ϲ�ȣ</th>
                                <th rowspan="2" scope="col" height="30" align="center">�����Һ����<br>���Ⱓ</th>
                                <th rowspan="2" scope="col" height="30" align="center">���Ա�<br>������</th>
                                <th colspan="3" scope="col" height="30" align="center" style="border-bottom:1px solid #e3e3e3;">������ ��ȯ��</th>
                                <th rowspan="2" scope="col" height="30" align="center">�����ݾ�(��)</th>
                              </tr>
                              <tr>
                                <th scope="col" height="30" align="center" style="border-left:1px solid #e3e3e3;">��</th>
                                <th scope="col" height="30" align="center">����</th>
                                <th scope="col" height="30" align="center">����</th>
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
                                <td height="30" align="center"><%=rs("s_type")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							end if
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>                    

                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
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
                                <td colspan="8" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">2) �Ӵ��� ��೻��</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">�Ӵ��μ���<br>(��ȣ)</th>
                                <th rowspan="2" scope="col" height="30" align="center">�ֹε�Ϲ�ȣ<br>(����ڹ�ȣ)</th>
                                <th rowspan="2" scope="col" height="30" align="center">��������</th>
                                <th rowspan="2" scope="col" height="30" align="center">���ð��<br>����(��)</th>
                                <th rowspan="2"scope="col" height="30" align="center">�Ӵ�����༭ �� �ּ���</th>
                                <th colspan="2" scope="col" height="30" align="center" style="border-bottom:1px solid #e3e3e3;">��༭��<br>�Ӵ��� ���Ⱓ</th>
                                <th rowspan="2" scope="col" height="30" align="center">����������(��)</th>
                              </tr>
                              <tr>
                                <th scope="col" height="30" align="center" style="border-left:1px solid #e3e3e3;">������</th>
                                <th scope="col" height="30" align="center">������</th>
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
                                <td height="30" align="center"><%=rs("s_type")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_name")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="left"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                            <tr>
                                <td colspan="8" height="30" align="left" scope="col" style="border-bottom:1px solid #e3e3e3;">�� �������� �����ڵ� - �ܵ�����:1, �ٰ���:2, �ټ�������:3, ��������:4, ����Ʈ:5, ���ǽ���:6, ��Ÿ:7<br><br>
                                �� ��༭�� �Ӵ������Ⱓ - �Խ��ϰ� �������� ���ÿ� ���� ����(����) 2013.01.01.</td>
                            </tr>
						</tbody>
					</table>   

                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="100%" >
                        </colgroup>
						 <thead>
                              <tr>
								<td scope="col" height="30" align="center" style=" border-bottom:2px solid #515254;">�� �� �� ��</td>
							  </tr>
                              <tr>
								<td scope="col" height="30" align="left" >
                                1. ������ ������ ������ �� ���������ڱ� ���ӱ� ������ ��ȯ�� ������ �޴� �ٷμҵ��ڿ� ���ؼ��� �ش� �ҵ������ ���� ���� �ۼ��ؾ� �մϴ�.<br><br>
                                2. �ش� �Ӵ��� ��ະ�� ���� �հ��� �����ס������ݻ�ȯ�װ� �����ݾ��� ������, �����ݾ��� 0�ΰ�쿡�� ���� �ʽ��ϴ�.<br><br>
                                3. ���������� �ܵ�����, �ٰ���, �ټ�������, ��������, ����Ʈ, ���ǽ���, ��Ÿ �߿��� �ش�Ǵ� �������� �����ڵ带 �����ϴ�.<br><br>
                                4. ������������ �����Ⱓ ������(12.31.) ������ ������������ �����ϴ�.</td>
                             </tr>
                        </thead>
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

