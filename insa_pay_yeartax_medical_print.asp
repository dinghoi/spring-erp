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

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�Ƿ�� ���޸���"
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
				<form action="insa_pay_yeartax_medical_print.asp" method="post" name="frm">
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
                              <td colspan="4" height="30" align="center" class="style12C">(<%=inc_yyyy%>) �� �Ƿ�� ���޸�</td>
						    </tr>
						</thead>
				  </table>
					<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <th colspan="4" height="30" align="center" scope="col" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                                <th colspan="3" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">����ó</th>
                                <th colspan="2" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">���޸�</th>
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
                              <tr>
                                <td colspan="7" class="first" height="30" align="center">�հ�</td>
                                <td align="right"><%=formatnumber(tot_cnt,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
							  </tr>                        
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
                                <td class="first" height="30" align="center"><%=m_rel%>&nbsp;</td>
                                <td align="center"><%=rs("m_national")%>&nbsp;</td>
                                <td align="center"><%=rs("m_person_no")%>&nbsp;</td>
                                <td align="center"><%=m_bon%>&nbsp;</td>
                                <td align="center"><%=rs("m_trade_no")%>&nbsp;</td>
                                <td align="center"><%=rs("m_trade_name")%>&nbsp;</td>
                                <td align="center"><%=m_data_gubun%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("m_cnt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("m_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
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

