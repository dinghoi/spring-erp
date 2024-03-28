<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

cost_month = request("cost_month")
sales_saupbu = request("sales_saupbu")
if sales_saupbu = "��Ÿ�����" then
	sales_saupbu = ""
end if

slip_month = mid(cost_month,1,4) + "-" + mid(cost_month,5,2)

title_line = cost_month + "�� " + sales_saupbu + " ��뼼�� ����"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

i = 0

sql = "select * from insure_per where insure_year = '"&mid(cost_month,1,4)&"'"
set rs_etc=dbconn.execute(sql)
insure_tot_per = rs_etc("insure_tot_per")
income_tax_per = rs_etc("income_tax_per")
annual_pay_per = rs_etc("annual_pay_per")
retire_pay_per = rs_etc("retire_pay_per")
rs_etc.close()

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��뱸��</th>
								<th scope="col">��������</th>
								<th scope="col">�������</th>
								<th scope="col">��翵�������</th>
								<th scope="col">��꼭����</th>
								<th scope="col">���ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��</th>
								<th scope="col">������</th>
								<th scope="col">����ó</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��������</th>
								<th scope="col">�������</th>
								<th scope="col">���־�ü</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">���೻��</th>
							</tr>
						</thead>
						<tbody>
						<%
						tot_tax = 0
						if (saupbu = sales_saupbu and position = "�������") or (saupbu = sales_saupbu and position = "������") or sales_grade = "0" then 
							if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
								sql = "select * from pay_month_give where pmg_id <>'4' and cost_center ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') ORDER BY pmg_id, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name, pmg_reside_place, pmg_reside_company, pmg_emp_name"
							  else	
								sql = "select * from pay_month_give where pmg_id <>'4' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') ORDER BY pmg_id, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name, pmg_reside_place, pmg_reside_company, pmg_emp_name"
							end if
							Rs.Open Sql, Dbconn, 1

							do until rs.eof
								tax_bill_yn = "�Ϲ�"
								gubun = "�ΰǺ�"
								account = "������"
								if rs("pmg_id") = "1" then
									account = "�޿�"
								end if
								if rs("pmg_id") = "2" then
									account = "��"
								end if
								if rs("pmg_id") = "3" then
									account = "��õ�μ�Ƽ��"
								end if
								cost_center  = rs("cost_center")
								mg_saupbu    = rs("mg_saupbu")
								emp_company  = rs("pmg_company")
								bonbu        = rs("pmg_bonbu")
								saupbu       = rs("pmg_saupbu")
								team         = rs("pmg_team")
								org_name     = rs("pmg_org_name")
								reside_place = rs("pmg_reside_place")
								company      = rs("pmg_reside_company")
								emp_name     = rs("pmg_emp_name")
								slip_date    = rs("pmg_yymm")
								slip_seq     = rs("pmg_emp_no")
								customer     = ""
								price        = rs("pmg_give_total")
								cost         = rs("pmg_give_total")
								cost_vat     = 0
								slip_memo    = ""
								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()

							if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
								sql = "select cost_center,pmg_company,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where pmg_id = '1' and cost_center ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') GROUP BY pmg_id, pmg_company ORDER BY pmg_company"
							  else	
								sql = "select cost_center,pmg_company,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where pmg_id = '1' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') GROUP BY pmg_id, cost_center,pmg_company ORDER BY pmg_company"
							end if
							Rs.Open Sql, Dbconn, 1

							do until rs.eof

                                'insure_tot = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * insure_tot_per / 100)	
                                insure_tot = clng((clng(rs("tot_cost"))) * insure_tot_per / 100)	
                                'income_tax = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * income_tax_per / 100)		
                                income_tax = clng((clng(rs("tot_cost"))) * income_tax_per / 100)		
								annual_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * annual_pay_per / 100)		
								retire_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * retire_pay_per / 100)		
								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>4�뺸���</td>
								<td><%=rs("cost_center")%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=rs("pmg_company")%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(insure_tot,0)%></td>
							  	<td class="right"><%=formatnumber(insure_tot,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>�ҵ漼��������</td>
								<td><%=rs("cost_center")%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=rs("pmg_company")%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(income_tax,0)%></td>
							  	<td class="right"><%=formatnumber(income_tax,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>��������</td>
								<td><%=rs("cost_center")%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=rs("pmg_company")%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(annual_pay,0)%></td>
							  	<td class="right"><%=formatnumber(annual_pay,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>��������</td>
								<td><%=rs("cost_center")%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=rs("pmg_company")%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(retire_pay,0)%></td>
							  	<td class="right"><%=formatnumber(retire_pay,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						end if
					
						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							sql = "select * from pay_alba_cost where cost_center ='"&sales_saupbu&"' and rever_yymm = '"&cost_month&"' ORDER BY cost_center,give_date,mg_saupbu,org_name, draft_man"
						  else
							sql = "select * from pay_alba_cost where mg_saupbu ='"&sales_saupbu&"' and (cost_center ='������' or cost_center ='����������') and rever_yymm = '"&cost_month&"' ORDER BY cost_center,give_date,mg_saupbu,org_name, draft_man"
						end if
						Rs.Open Sql, Dbconn, 1
						do until rs.eof
						  	tax_bill_yn  = "�Ϲ�"
							gubun        = "�ΰǺ�"
							account      = "�˹ٺ�"
							cost_center  = rs("cost_center")
							mg_saupbu    = rs("mg_saupbu")
							emp_company  = rs("company")
							bonbu        = rs("bonbu")
							saupbu       = rs("saupbu")
							team         = rs("team")
							org_name     = rs("org_name")
							reside_place = ""
							company      = rs("cost_company")
							emp_name     = rs("draft_man")
							slip_date    = rs("give_date")
							slip_seq     = rs("draft_no")
							customer     = ""
							price        = rs("alba_give_total")
							cost         = rs("alba_give_total")
							cost_vat     = 0
							slip_memo    = rs("draft_tax_id")
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							sql = "select * from general_cost where (pl_yn = 'Y') and cancel_yn ='N' and cost_center ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
						  else	
							sql = "select * from general_cost where (pl_yn = 'Y') and cancel_yn ='N' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
						end if
						Rs.Open Sql, Dbconn, 1
						do until rs.eof
							if rs("tax_bill_yn") = "Y" then
								tax_bill_yn = "���ݰ�꼭" 
							  else
							  	tax_bill_yn = "�Ϲ�"
							end if
							gubun        = rs("slip_gubun")
							account      = rs("account")
							cost_center  = rs("cost_center")
							mg_saupbu    = rs("mg_saupbu")
							emp_company  = rs("emp_company")
							bonbu        = rs("bonbu")
							saupbu       = rs("saupbu")
							team         = rs("team")
							org_name     = rs("org_name")
							reside_place = rs("reside_place")
							company      = rs("company")
							emp_name     = rs("emp_name")
							slip_date    = rs("slip_date")
							slip_seq     = rs("slip_seq")
							customer     = rs("customer")
							price        = rs("price")
							cost         = rs("cost")
							cost_vat     = rs("cost_vat")
							slip_memo    = rs("slip_memo")
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							sql = "select * from transit_cost where cancel_yn ='N' and cost_center ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
						  else
							sql = "select * from transit_cost where cancel_yn ='N' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
						end if
						Rs.Open Sql, Dbconn, 1
						do until rs.eof
						  	tax_bill_yn  = "�Ϲ�"
							gubun        = "�����"
							account      = rs("car_owner")
							cost_center  = rs("cost_center")
							mg_saupbu    = rs("mg_saupbu")
							emp_company  = rs("emp_company")
							bonbu        = rs("bonbu")
							saupbu       = rs("saupbu")
							team         = rs("team")
							org_name     = rs("org_name")
							reside_place = rs("reside_place")
							company      = rs("company")
							emp_name     = rs("user_name")
							slip_date    = rs("run_date")
							slip_seq     = rs("run_seq")
							customer     = ""
							price        = rs("somopum") + rs("oil_price") + rs("fare") + rs("parking") + rs("toll")
							cost         = rs("somopum") + rs("oil_price") + rs("fare") + rs("parking") + rs("toll")
							cost_vat     = 0
							slip_memo    = rs("run_memo")
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							sql = "select * from transit_cost where cancel_yn ='N' and repair_cost > 0 and cost_center ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
						  else
							sql = "select * from transit_cost where cancel_yn ='N' and (cost_center ='������' or cost_center ='����������') and repair_cost > 0 and mg_saupbu ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
						end if
						Rs.Open Sql, Dbconn, 1
						do until rs.eof
						  	tax_bill_yn  = "�Ϲ�"
							gubun        = "�����"
							account      = "����������"
							cost_center  = rs("cost_center")
							mg_saupbu    = rs("mg_saupbu")
							emp_company  = rs("emp_company")
							bonbu        = rs("bonbu")
							saupbu       = rs("saupbu")
							team         = rs("team")
							org_name     = rs("org_name")
							reside_place = rs("reside_place")
							company      = rs("company")
							emp_name     = rs("user_name")
							slip_date    = rs("run_date")
							slip_seq     = rs("run_seq")
							customer     = ""
							price        = rs("repair_cost")
							cost         = rs("repair_cost")
							cost_vat     = 0
							slip_memo    = rs("run_memo")
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							sql = "select * from card_slip where (pl_yn = 'Y') and (card_type not like '%����%' or com_drv_yn = 'Y') and cost_center ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
						  else
							sql = "select * from card_slip where (pl_yn = 'Y') and (card_type not like '%����%' or com_drv_yn = 'Y') and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
						end If
						
						'// and mg_saupbu ='"&sales_saupbu&"' and
						'//where (pl_yn = 'Y') and (card_type not like '%����%' or com_drv_yn = 'Y') and
						Rs.Open Sql, Dbconn, 1
						do until rs.eof
						  	tax_bill_yn   = "�Ϲ�"
							gubun         = "����ī��"
							account       = rs("account")
							cost_center   = rs("cost_center")
							mg_saupbu     = rs("mg_saupbu")
							emp_company   = rs("emp_company")
							bonbu         = rs("bonbu")
							saupbu        = rs("saupbu")
							team          = rs("team")
							org_name      = rs("org_name")
							reside_place  = rs("reside_place")
							company       = rs("reside_company")
							emp_name      = rs("emp_name")
							slip_date     = rs("slip_date")
							slip_seq      = rs("approve_no")
							customer      = rs("customer")
							price         = rs("price")
							cost          = rs("cost")
							cost_vat      = rs("cost_vat")
							slip_memo     = rs("account_item")
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

