<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

car_no=Request("car_no")
from_date=Request("from_date")
to_date=Request("to_date")
view_condi = car_no

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_drv = Server.CreateObject("ADODB.Recordset")
Set Rs_insu = Server.CreateObject("ADODB.Recordset")
Set Rs_pen = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from car_info where car_no = '"&view_condi&"'"
Rs.Open Sql, Dbconn, 1


title_line = "���� ��Ż ���� ��Ȳ"

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
                              <col width="10%" >
							  <col width="*" >
                              <col width="8%" >
							  <col width="6%" >
							  <col width="6%" >
							  <col width="12%" >
							  <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
							  <col width="12%" >
							  <col width="6%" >
							  <col width="6%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <th class="first" scope="col" height="30" align="center">������ȣ</th>
								<th scope="col" height="30" align="center">����</th>
                                <th scope="col" height="30" align="center">����</th>
								<th scope="col" height="30" align="center">����</th>
								<th scope="col" height="30" align="center">����</th>
								<th scope="col" height="30" align="center">����<br>����</th>
								<th scope="col" height="30" align="center">���������</th>
                                <th scope="col" height="30" align="center">�����</th>
                                <th scope="col" height="30" align="center">����Ⱓ</th>
								<th scope="col" height="30" align="center">������</th>
								<th scope="col" height="30" align="center">����KM</th>
								<th scope="col" height="30" align="center">�����˻���</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof
                             
							 owner_emp_name = ""
							 owner_emp_no = rs("owner_emp_no")
						     if rs("owner_emp_name") = "" or isnull(rs("owner_emp_name")) then
							     Sql="select * from emp_master where emp_no = '"&owner_emp_no&"'"
	                             Set rs_emp=DbConn.Execute(Sql)
								 owner_emp_name = rs_emp("emp_name")
							   else 
							     owner_emp_name = rs("owner_emp_name")
							 end if
							if rs("last_check_date") = "1900-01-01"  then
	                               last_check_date = ""
							   else 
							       last_check_date = rs("last_check_date")
	                        end if
	                        if rs("end_date") = "1900-01-01" then
	                               end_date = ""
							   else 
							       end_date = rs("end_date")
	                        end if
							if rs("car_year") = "1900-01-01" then
	                               car_year = ""
							   else 
							       car_year = rs("car_year")
	                        end if
	           		 %>
							<tr>
                                <td height="30" align="center"><%=rs("car_no")%>&nbsp;</td>
								<td align="center"><%=rs("car_name")%></td>
                                <td align="center"><%=car_year%>&nbsp;</td>
								<td align="center"><%=rs("oil_kind")%></td>
								<td align="center"><%=rs("car_owner")%></td>
								<td align="center"><%=rs("buy_gubun")%>&nbsp;<%=rs("rental_company")%></td>
								<td align="center"><%=rs("car_reg_date")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("insurance_amt"),0)%>&nbsp;</td>
                                <td align="center"><%=rs("insurance_date")%>&nbsp;</td>
                                <td align="center"><%=owner_emp_name%>(<%=rs("owner_emp_no")%>)&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("last_km"),0)%>&nbsp;</td>
								<td align="center"><%=last_check_date%>&nbsp;</td>
							</tr>
					<%
							rs.movenext()
						loop
						rs.close()
						
					%>
						</tbody>
					</table>
         <%
						sql = "select * from car_insurance where ins_car_no = '"&view_condi&"' ORDER BY ins_car_no,ins_date ASC"
                        Rs_insu.Open Sql, Dbconn, 1
						if not Rs_insu.EOF or not Rs_insu.BOF then
		 %>                    
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="7%" >
                              <col width="6%" >
							  <col width="10%" >
                              <col width="6%" >
                              <col width="7%" >
                              <col width="7%" >
                              <col width="6%" >
                              <col width="6%" >
                              <col width="6%" >
                              <col width="6%" >
                              <col width="7%" >
                              <col width="6%" >
                              <col width="4%" >
                              <col width="*" >

                        </colgroup>
						 <thead>
                              <tr>
                                <th class="first" scope="col" height="30" align="center">������ȣ</th>
                                <th scope="col" align="center">������</th>
                                <th scope="col" align="center">�����</th>
                                <th scope="col" align="center">����Ⱓ</th>
                                <th scope="col" align="center">�����</th>
                                <th scope="col" align="center">����1</th>
                                <th scope="col" align="center">����2</th>
                                <th scope="col" align="center">�빰</th>
                                <th scope="col" align="center">�ڱ⺸��</th>
                                <th scope="col" align="center">������</th>
                                <th scope="col" align="center">����</th>
                                <th scope="col" align="center">����</th>
                                <th scope="col" align="center">���<br>�⵿</th>
                                <th scope="col" align="center">��೻��</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
				
						do until Rs_insu.eof
                             
							  car_no = Rs_insu("ins_car_no")
							  
							  Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
                              Set rs_car = DbConn.Execute(SQL)
							  if not rs_car.eof then
									car_name = rs_car("car_name")
									car_year = rs_car("car_year")
									car_reg_date = rs_car("car_reg_date")
	                             else
								    car_name = ""
									car_year = ""
									car_reg_date = ""
                              end if
                              rs_car.close()
	           			%>
							<tr>
                                <td height="30" align="center"><%=Rs_insu("ins_car_no")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_date")%>&nbsp;</td>
								<td align="center"><%=Rs_insu("ins_company")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_last_date")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs_insu("ins_amount"),0)%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_man1")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_man2")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_object")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_self")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_injury")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_self_car")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_age")%>&nbsp;</td>
                                <td align="center"><%=Rs_insu("ins_scramble")%>&nbsp;</td>
                         <% if Rs_insu("ins_contract_yn") = "Y" then %>
                                <td align="left">��೻������&nbsp;</td>
                         <%    else %>
                                <td align="left">��೻�������(<%=Rs_insu("ins_comment")%>)&nbsp;</td>
                         <% end if %>
							</tr>
						<%
							Rs_insu.movenext()
						loop
						%>
						</tbody>
					</table>                    
         <% 
		                Rs_insu.close()
			  end if %>
         <%
						tot_fare = 0
                        tot_oil_price = 0
						tot_parking = 0
                        tot_toll = 0
                        sql = "select * from transit_cost where car_no = '"&view_condi&"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' ORDER BY car_no,run_date,run_seq ASC"
						Rs.Open Sql, Dbconn, 1
                        do until rs.eof
                              tot_fare = tot_fare + int(rs("fare"))
	                          tot_oil_price = tot_oil_price + int(rs("oil_price"))
							  tot_parking = tot_parking + int(rs("parking"))
							  tot_toll = tot_toll + int(rs("toll"))
	                       rs.movenext()
                        loop
                        rs.close()	
						
						sql = "select * from transit_cost where car_no = '"&view_condi&"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' ORDER BY car_no,run_date,run_seq ASC"
                        Rs_drv.Open Sql, Dbconn, 1
						if not Rs_drv.EOF or not Rs_drv.BOF then
		 %>          
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="6%" >
                              <col width="6%" >
							  <col width="5%" >
							  <col width="5%" >
							  <col width="4%" >
							  <col width="8%" >
							  <col width="9%" >
							  <col width="5%" >
							  <col width="8%" >
							  <col width="*" >
							  <col width="5%" >
							  <col width="6%" >
							  <col width="5%" >
							  <col width="5%" >
							  <col width="4%" >
							  <col width="4%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">������ȣ</th>
                                <th rowspan="2" scope="col" height="30" align="center">��������</th>
								<th rowspan="2" scope="col" height="30" align="center">������</th>
								<th rowspan="2" scope="col" height="30" align="center">����</th>
								<th rowspan="2" scope="col" height="30" align="center">����<br>/<br>����<br>����</th>
								<th colspan="3" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">�� ��</th>
								<th colspan="3" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">�� ��</th>
								<th rowspan="2" scope="col" height="30" align="center">�������</th>
								<th colspan="4" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">�� �� </th>
							</tr>
							<tr>
								<th scope="col" height="30" align="center" style=" border-left:1px solid #e3e3e3;">��ü��</th>
								<th scope="col" height="30" align="center">�����</th>
								<th scope="col" height="30" align="center">���KM</th>
								<th scope="col" height="30" align="center">��ü��</th>
								<th scope="col" height="30" align="center">������</th>
								<th scope="col" height="30" align="center">����KM</th>
								<th scope="col" height="30" align="center">���߱���</th>
								<th scope="col" height="30" align="center">�����ݾ�</th>
								<th scope="col" height="30" align="center">������</th>
								<th scope="col" height="30" align="center">�����</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
					
						do until Rs_drv.eof
						    emp_no = Rs_drv("mg_ce_id")
							Sql = "select * from emp_master where emp_no = '"+emp_no+"'"
	                        Set Rs_emp = DbConn.Execute(SQL)
	                        if not Rs_emp.EOF or not Rs_emp.BOF then
			                       drv_owner_emp_name = rs_emp("emp_name")
                               else
                                   drv_owner_emp_name = emp_no
							end if
							
							if Rs_drv("start_km") = "" or isnull(Rs_drv("start_km")) then
								start_view = 0
							  else
							  	start_view = Rs_drv("start_km")
							end if
							if Rs_drv("end_km") = "" or isnull(Rs_drv("end_km")) then
								end_view = 0
							  else
							  	end_view = Rs_drv("end_km")
							end if
							run_km = Rs_drv("far")
                             
	           			%>
							<tr>
                                <td class="first" height="30" align="center"><%=Rs_drv("car_no")%></td>
                                <td height="30" align="center"><%=Rs_drv("run_date")%></td>
								<td height="30" align="center"><%=drv_owner_emp_name%></td>
								<td height="30" align="center"><%=Rs_drv("car_owner")%></td>
								<td height="30" align="center">
								<% if Rs_drv("car_owner") = "���߱���" then %>
								       <%=Rs_drv("transit")%>
								<%   else	%>                                
								       <%=Rs_drv("oil_kind")%>
								<% end if %>
                                </td>
								<td height="30" align="center"><%=Rs_drv("start_company")%>&nbsp;</td>
								<td align="left"><%=Rs_drv("start_point")%></td>
								<td align="right"><%=formatnumber(start_view,0)%></td>
								<td height="30" align="center"><%=Rs_drv("end_company")%>&nbsp;</td>
								<td align="left"><%=Rs_drv("end_point")%></td>
								<td align="right"><%=formatnumber(end_view,0)%></td>
								<td height="30" align="center"><%=Rs_drv("run_memo")%></td>
								<td align="right"><%=formatnumber(Rs_drv("fare"),0)%></td>
								<td align="right"><%=formatnumber(Rs_drv("oil_price"),0)%></td>
								<td align="right"><%=formatnumber(Rs_drv("parking"),0)%></td>
								<td align="right"><%=formatnumber(Rs_drv("toll"),0)%></td>
							</tr>
						<%
							Rs_drv.movenext()
						loop
						%>
                            <tr>
								<td colspan="12" height="30" align="center" style="background:#ffe8e8;">�Ѱ�</td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_fare,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_oil_price,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_parking,0)%></td>
                                <td align="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_toll,0)%></td>
							</tr>                                
						</tbody>
					</table>   
         <% 
		                Rs_drv.close()
			  end if %>
         <%
						tot_amount = 0
                        sql = "select * from car_as where as_car_no = '"&view_condi&"' and as_date >= '"+from_date+"' and as_date <= '"+to_date+"' ORDER BY as_car_no,as_date,as_seq ASC"
						Rs.Open Sql, Dbconn, 1
                        do until rs.eof
                              tot_amount = tot_amount + int(rs("as_amount"))
	                       rs.movenext()
                        loop
                        rs.close()							
						
						sql = "select * from car_as where as_car_no = '"&view_condi&"' and as_date >= '"+from_date+"' and as_date <= '"+to_date+"' ORDER BY as_car_no,as_date,as_seq ASC"
                        Rs_as.Open Sql, Dbconn, 1
						if not Rs_as.EOF or not Rs_as.BOF then
		 %>          
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="8%" >
                              <col width="10%" >
							  <col width="12%" >
                              <col width="8%" >
							  <col width="15%" >
							  <col width="*" >
							  <col width="8%" >
                              <col width="6%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <th class="first" scope="col" height="30" align="center">������ȣ</th>
                                <th scope="col" align="center">����</th>
								<th scope="col" align="center">������</th>
                                <th scope="col" align="center">AS����</th>
								<th scope="col" align="center">AS����</th>
								<th scope="col" align="center">��������</th>
								<th scope="col" align="center">�������</th>
                                <th scope="col" align="center">���籸��</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
					
						do until Rs_as.eof

	           		 %>
							<tr>
                                <td class="first" height="30" align="center"><%=Rs_as("as_car_no")%></td>
                                <td align="center"><%=Rs_as("as_car_name")%></td>
                                <td align="center"><%=Rs_as("as_owner_emp_name")%>(<%=Rs_as("as_owner_emp_no")%>)</td>
                                <td align="center"><%=Rs_as("as_date")%></td>
								<td align="left"><%=Rs_as("as_cause")%></td>
								<td align="left"><%=Rs_as("as_solution")%></td>
                                <td align="right"><%=formatnumber(Rs_as("as_amount"),0)%></td>
                                <td align="center"><%=Rs_as("as_amount_sign")%></td>
							</tr>
						<%
							Rs_as.movenext()
						loop
						%>
                            <tr>
								<td colspan="6" height="30" align="center" style="background:#ffe8e8;">�Ѱ�</td>
                                <td align="right" style="font-size:12px; background:#ffe8e8;"><%=formatnumber(tot_amount,0)%>&nbsp;</td>
                                <td height="30" align="center" style="background:#ffe8e8;">&nbsp;</td>
							</tr>                        
						</tbody>
					</table> 
         <% 
		                Rs_as.close()
			  end if %>                    
         <%
						tot_amount = 0
                        tot_in_amt = 0
                        sql = "select * from car_penalty where pe_car_no = '"&view_condi&"' and pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' ORDER BY pe_car_no,pe_date,pe_seq ASC"
						Rs.Open Sql, Dbconn, 1
                        do until rs.eof
                              tot_amount = tot_amount + int(rs("pe_amount"))
	                          tot_in_amt = tot_in_amt + int(rs("pe_in_amt"))
	                       rs.movenext()
                        loop
                        rs.close()	
						jan_amount = tot_amount - tot_in_amt
						
						sql = "select * from car_penalty where pe_car_no = '"&view_condi&"' and pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' ORDER BY pe_car_no,pe_date,pe_seq ASC"
                        Rs_pen.Open Sql, Dbconn, 1
						if not Rs_pen.EOF or not Rs_pen.BOF then
		 %>          
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="7%" >
                              <col width="6%" >
							  <col width="8%" >
							  <col width="8%" >
                              <col width="6%" >
							  <col width="8%" >
                              <col width="6%" >
							  <col width="*" >
							  <col width="6%" >
                              <col width="6%" >
                              <col width="8%" >
                              <col width="6%" >
                              <col width="8%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <th class="first" scope="col" height="30" align="center">������ȣ</th>
                                <th scope="col" align="center">����</th>
								<th scope="col" align="center">������</th>
								<th scope="col" align="center">�μ�</th>
                                <th scope="col" align="center">��������</th>
								<th scope="col" align="center">���ݳ���</th>
								<th scope="col" align="center">���·�</th>
								<th scope="col" align="center">�������</th>
                                <th scope="col" align="center">��������</th>
                                <th scope="col" align="center">�뺸����</th>
                                <th scope="col" align="center">�뺸���</th>
                                <th scope="col" align="center">�̳�</th>
                                <th scope="col" align="center">���</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
					
						do until Rs_pen.eof
						
						  car_no = Rs_pen("pe_car_no")
						  if Rs_pen("pe_in_date") = "1900-01-01"  then
	                               pe_in_date = ""
							   else 
							       pe_in_date = Rs_pen("pe_in_date")
	                       end if
	                       if Rs_pen("pe_notice_date") = "1900-01-01" then
	                               pe_notice_date = ""
							   else 
							       pe_notice_date = Rs_pen("pe_notice_date")
	                       end if
							  
		                   Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
                           Set rs_car = DbConn.Execute(SQL)
		                   if not rs_car.eof then
		                        	car_name = rs_car("car_name")
		                    		car_year = rs_car("car_year")
			                    	car_reg_date = rs_car("car_reg_date")
		                    		car_use_dept = rs_car("car_use_dept")
	                    			car_company = rs_car("car_company")
	                     			car_use = rs_car("car_use")
									car_owner = rs_car("car_owner")
	                    			owner_emp_name = rs_car("owner_emp_name")
	                    			owner_emp_no = rs_car("owner_emp_no")
	                     			oil_kind = rs_car("oil_kind")
	                          else
	                     		    car_name = ""
	                    			car_year = ""
			                      	car_reg_date = ""
			                    	car_use_dept = ""
		                    		car_company = ""
		                    		car_use = ""
									car_owner = ""
	                    			owner_emp_name = ""
		                    		owner_emp_no = ""
	                    			oil_kind = ""
                           end if
                           rs_car.close()

	           		 %>
							<tr>
                                <td height="30" align="center"><%=Rs_pen("pe_car_no")%>&nbsp;</td>
                                <td align="center"><%=car_name%>&nbsp;</td>
                                <td align="center"><%=owner_emp_name%>(<%=owner_emp_no%>)&nbsp;</td>
                                <td align="center"><%=car_use_dept%>&nbsp;</td>
                                <td align="center"><%=Rs_pen("pe_date")%>&nbsp;</td>
								<td align="left"><%=Rs_pen("pe_comment")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(Rs_pen("pe_amount"),0)%>&nbsp;</td>
                                <td align="left"><%=Rs_pen("pe_place")%>&nbsp;</td>
                                <td align="center"><%=pe_in_date%>&nbsp;</td>
                                <td align="center"><%=pe_notice_date%>&nbsp;</td>
                                <td align="left"><%=Rs_pen("pe_notice")%>&nbsp;</td>
                                <td align="left"><%=Rs_pen("pe_default")%>&nbsp;</td>
                                <td align="left"><%=Rs_pen("pe_bigo")%>&nbsp;</td>
							</tr>
						<%
							Rs_pen.movenext()
						loop
						%>
                            <tr>
								<td colspan="4" height="30" align="center" style="background:#ffe8e8;">�Ѱ�</td>
                                <td height="30" align="center" style="background:#ffe8e8;">���·� ��</td>
                                <td colspan="2" align="right" style="font-size:12px; background:#ffe8e8;"><%=formatnumber(tot_amount,0)%>&nbsp;</td>
                                <td height="30" align="center" style="background:#ffe8e8;">���� ��</td>
                                <td colspan="2" align="right" style="font-size:12px; background:#ffe8e8;"><%=formatnumber(tot_in_amt,0)%>&nbsp;</td>
                                <td height="30" align="center" style="background:#ffe8e8;">�̳� ��</td>
                                <td colspan="2" align="right" style="font-size:12px; background:#ffe8e8;"><%=formatnumber(jan_amount,0)%>&nbsp;</td>
							</tr>                        
						</tbody>
					</table> 
         <% 
		                Rs_pen.close()
			  end if %>      

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

